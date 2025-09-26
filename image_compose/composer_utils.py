
import io
from pathlib import Path
from PIL import Image, ImageFilter

# ------------------- 설정값 정의 -------------------
# 그림자 프리셋: UI와 CLI에서 공통으로 사용
SHADOW_PRESETS = {
    "off":    {"blur": 0,  "alpha": 0,   "offset_x": 0.0, "offset_y": 0.0},
    "light":  {"blur": 6,  "alpha": 100, "offset_x": 0.006, "offset_y": 0.006},
    "medium": {"blur": 14, "alpha": 160, "offset_x": 0.012, "offset_y": 0.012},
    "strong": {"blur": 24, "alpha": 220, "offset_x": 0.018, "offset_y": 0.018},
}

# ------------------- 유틸리티 함수 -------------------
def ensure_rgba(img: Image.Image) -> Image.Image:
    """이미지를 RGBA 모드로 변환합니다."""
    if img.mode == "RGBA":
        return img
    if img.mode in ("LA", "P"):
        return img.convert("RGBA")
    if img.mode == "RGB":
        # RGB는 완전 불투명 알파를 추가
        rgba = Image.new("RGBA", img.size, (0, 0, 0, 0))
        rgba.paste(img, (0, 0))
        return rgba
    return img.convert("RGBA")

def has_useful_alpha(img: Image.Image) -> bool:
    """알파가 의미 있게 존재하는지 확인 (완전 불투명/완전 투명만 있는 경우 False)."""
    img = ensure_rgba(img)
    a = img.getchannel("A")
    extrema = a.getextrema()  # (min, max)
    if not extrema:
        return False
    min_a, max_a = extrema
    return not (min_a == 255 and max_a == 255) and not (min_a == 0 and max_a == 0)

def load_images_from_folder(folder: Path):
    """폴더에서 지원 포맷 이미지를 로드 (경로 목록 반환)."""
    exts = {".png", ".jpg", ".jpeg", ".webp"}
    return [p for p in sorted(Path(folder).glob("*")) if p.suffix.lower() in exts]

def compute_anchor_position(bg_size, fg_size, anchor: str):
    """기준점에 맞춰 아이템이 배치될 좌상단(x, y) 좌표를 계산합니다."""
    W, H = bg_size
    w, h = fg_size
    positions = {
        "center": ((W - w) // 2, (H - h) // 2),
        "top": ((W - w) // 2, 0),
        "bottom": ((W - w) // 2, H - h),
        "left": (0, (H - h) // 2),
        "right": (W - w, (H - h) // 2),
        "top-left": (0, 0),
        "top-right": (W - w, 0),
        "bottom-left": (0, H - h),
        "bottom-right": (W - w, H - h),
    }
    return positions.get(anchor, positions["center"])

# ------------------- 핵심 합성 함수 -------------------
def compose_one_bytes(item_img: Image.Image, template_img: Image.Image, **opts) -> tuple | None:
    """
    하나의 템플릿에 하나의 아이템 이미지를 합성하여 BytesIO 객체로 반환합니다.
    opts:
      - anchor: str (default "center")
      - resize_ratio: float (default 1.0)  # 아이템 크기에 대한 배율
      - shadow_preset: str in SHADOW_PRESETS (default "off")
      - out_format: "JPEG" | "PNG" (default "JPEG")
      - quality: int (default 92) for JPEG
    """
    # 0) 입력 이미지 보정
    item_rgba = ensure_rgba(item_img)
    template_rgba = ensure_rgba(template_img)

    # 1) 아이템 리사이즈
    ratio = float(opts.get("resize_ratio", 1.0))
    if ratio <= 0:
        ratio = 1.0
    if ratio != 1.0:
        new_size = (max(1, int(item_rgba.width * ratio)), max(1, int(item_rgba.height * ratio)))
        item_rgba = item_rgba.resize(new_size, Image.LANCZOS)

    # 2) 배치 좌표 계산
    anchor = opts.get("anchor", "center")
    x, y = compute_anchor_position(template_rgba.size, item_rgba.size, anchor)

    # 3) 최종 캔버스 준비 (템플릿 복사본)
    final_img = template_rgba.copy()

    # 4) 그림자 프리셋 적용 (항상 아이템 뒤에 먼저 합성)
    preset_name = str(opts.get("shadow_preset", "off"))
    preset = SHADOW_PRESETS.get(preset_name, SHADOW_PRESETS["off"])

    if preset.get("alpha", 0) > 0:
        # (a) 아이템 알파 추출
        alpha_mask = item_rgba.getchannel("A")

        # (b) 블러
        blur_radius = int(preset.get("blur", 0))
        if blur_radius > 0:
            alpha_blurred = alpha_mask.filter(ImageFilter.GaussianBlur(blur_radius))
        else:
            alpha_blurred = alpha_mask

        # (c) 강도 스케일 (0~255)
        scale = max(0, min(255, int(preset.get("alpha", 0)))) / 255.0
        alpha_scaled = alpha_blurred.point(lambda p: int(p * scale))

        # (d) 검정 RGB + 스케일된 알파로 그림자 RGBA 생성
        #     RGB는 0(검정), 알파는 alpha_scaled
        shadow_rgba = Image.new("RGBA", item_rgba.size, (0, 0, 0, 0))
        shadow_rgba.putalpha(alpha_scaled)

        # (e) 오프셋 계산 (템플릿 크기 기준, 기존 로직 유지)
        dx = int(template_rgba.width * float(preset.get("offset_x", 0.0)))
        dy = int(template_rgba.height * float(preset.get("offset_y", 0.0)))

        # (f) 최종 캔버스에 그림자 합성 (항상 먼저)
        shadow_layer = Image.new("RGBA", final_img.size, (0, 0, 0, 0))
        shadow_layer.paste(shadow_rgba, (x + dx, y + dy), shadow_rgba)
        final_img = Image.alpha_composite(final_img, shadow_layer)

    # 5) 아이템 합성 (항상 그림자 위에)
    item_layer = Image.new("RGBA", final_img.size, (0, 0, 0, 0))
    item_layer.paste(item_rgba, (x, y), item_rgba)
    final_img = Image.alpha_composite(final_img, item_layer)

    # 6) 저장
    img_buf = io.BytesIO()
    out_format = str(opts.get("out_format", "JPEG")).upper()
    if out_format == "JPEG":
        final_img = final_img.convert("RGB")
        final_img.save(img_buf, format="JPEG", quality=int(opts.get("quality", 92)))
        ext = "jpg"
    else:
        final_img.save(img_buf, format="PNG")
        ext = "png"
    img_buf.seek(0)
    return img_buf, ext

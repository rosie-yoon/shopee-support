import argparse
import zipfile
from pathlib import Path
from tqdm import tqdm
from PIL import Image
from composer_utils import compose_one_bytes, load_images_from_folder, SHADOW_PRESETS

def main(args):
    """CLI를 통해 배치 합성을 수행하는 메인 함수"""
    item_folder = Path(args.item_folder)
    template_folder = Path(args.template_folder)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    item_files = load_images_from_folder(item_folder)
    template_files = load_images_from_folder(template_folder)

    if not item_files:
        print(f"❌ Error: Item 폴더 '{item_folder}'에 이미지가 없습니다.")
        return
    if not template_files:
        print(f"❌ Error: Template 폴더 '{template_folder}'에 이미지가 없습니다.")
        return

    opts = {
        "anchor": args.anchor, "resize_ratio": args.resize_ratio,
        "shadow_preset": args.shadow_preset, "out_format": args.out_format,
        "quality": args.quality, "skip_if_no_alpha": True,
    }
    
    zip_path = out_dir.with_suffix(".zip")
    generated_count = 0

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        total_jobs = len(item_files) * len(template_files)
        with tqdm(total=total_jobs, desc="이미지 합성 중") as pbar:
            for item_name, item_path in item_files:
                item_img = Image.open(item_path)
                if not opts.get("skip_if_no_alpha") or has_useful_alpha(ensure_rgba(item_img)):
                    for template_name, template_path in template_files:
                        template_img = Image.open(template_path)
                        result = compose_one_bytes(item_img, template_img, **opts)
                        
                        if result:
                            img_buf, ext = result
                            shop_var = args.custom_variable if args.custom_variable else template_name
                            filename = f"{item_name}_C_{shop_var}.{ext}"
                            
                            save_path = out_dir / filename
                            save_path.write_bytes(img_buf.getvalue())
                            zf.write(save_path, arcname=filename)
                            generated_count += 1
                pbar.update(len(template_files))

    if generated_count > 0:
        print(f"✅ 완료! 총 {generated_count}개 이미지 생성.")
        print(f"   - 개별 파일: '{out_dir}'\n   - 압축 파일: '{zip_path}'")
    else:
        print("⚠️ 생성된 이미지가 없습니다. Item 파일에 투명 배경이 있는지 확인해주세요.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Thumb Craft - CLI 이미지 합성 도구")
    parser.add_argument("--item_folder", required=True, help="Item 이미지(투명 PNG/WEBP)가 있는 폴더")
    parser.add_argument("--template_folder", required=True, help="Template 이미지가 있는 폴더")
    parser.add_argument("--out_dir", default="C_out", help="결과물이 저장될 폴더")
    
    parser.add_argument("--anchor", default="center", choices=["center","top","bottom","left","right","top-left","top-right","bottom-left","bottom-right"], help="Item 배치 위치")
    parser.add_argument("--resize_ratio", type=float, default=1.0, help="Item 리사이즈 비율 (예: 0.8)")
    parser.add_argument("--shadow_preset", default="off", choices=SHADOW_PRESETS.keys(), help="적용할 그림자 프리셋")
    parser.add_argument("--out_format", default="JPEG", choices=["JPEG", "PNG"], help="출력 포맷")
    parser.add_argument("--quality", type=int, default=92, help="JPEG 품질 (1-100)")
    parser.add_argument("--custom_variable", default="", help="파일명에 사용할 Shop 구분값")

    from composer_utils import has_useful_alpha, ensure_rgba
    main(parser.parse_args())


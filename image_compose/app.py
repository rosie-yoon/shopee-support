from __future__ import annotations
from pathlib import Path
import io
import zipfile

import streamlit as st
from PIL import Image
import numpy as np

# 상대 임포트: image_compose/composer_utils.py 가 같은 폴더에 있어야 합니다.
# 패키지 인식을 위해 image_compose/__init__.py 도 반드시 존재해야 합니다.
from .composer_utils import compose_one_bytes, SHADOW_PRESETS, has_useful_alpha, ensure_rgba

BASE_DIR = Path(__file__).resolve().parent  # 필요 시 사용


def run():
    st.title("🎆 Cover Image")

    # ---- 세션 상태 초기화 ----
    def init_state():
        defaults = {
            "anchor": "center",
            "resize_ratio": 1.0,
            "shadow_preset": "off",
            "item_uploader_key": 0,
            "template_uploader_key": 0,
            "preview_img_bytes": None, # PIL 객체 대신 bytes를 저장하여 안정성 확보
            "download_info": None,
            "preview_index": 0,  # 현재 미리보기 아이템 인덱스
        }
        for k, v in defaults.items():
            if k not in st.session_state:
                st.session_state[k] = v

    init_state()
    ss = st.session_state

    # ---- 합성 미리보기 ----
    def update_preview(item_files, template_files, index):
        ss.preview_img_bytes = None
        # 유효한 인덱스이고 파일이 있는지 확인
        if not item_files or not template_files or index >= len(item_files):
            return

        item_img = Image.open(item_files[index]) # 지정된 인덱스의 아이템 사용
        template_img = Image.open(template_files[0]) # 템플릿은 항상 첫번째 것을 사용

        if not has_useful_alpha(ensure_rgba(item_img)):
            try:
                st.toast("투명 배경이 아닌 Item은 생성에서 제외됩니다.", icon="⚠️")
            except Exception:
                st.warning("투명 배경이 아닌 Item은 생성에서 제외됩니다.")
            return

        opts = {
            "anchor": ss.anchor,
            "resize_ratio": ss.resize_ratio,
            "shadow_preset": ss.shadow_preset,
            "out_format": "PNG",
        }
        result = compose_one_bytes(item_img, template_img, **opts)
        if result:
            buf, ext = result
            # PIL 객체가 아닌, raw bytes를 세션에 직접 저장
            ss.preview_img_bytes = buf.getvalue()

    # ---- 배치 합성 & Zip 생성 ----
    def run_batch_composition(item_files, template_files, fmt, quality, shop_variable):
        zip_buf = io.BytesIO()
        count = 0
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for item_file in item_files:
                item_img = Image.open(item_file)
                if not has_useful_alpha(ensure_rgba(item_img)):
                    continue

                for template_file in template_files:
                    template_img = Image.open(template_file)
                    opts = {
                        "anchor": ss.anchor,
                        "resize_ratio": ss.resize_ratio,
                        "shadow_preset": ss.shadow_preset,
                        "out_format": fmt,
                        "quality": quality,
                    }
                    result = compose_one_bytes(item_img, template_img, **opts)
                    if result:
                        img_buf, ext = result
                        item_name = Path(item_file.name).stem
                        shop_var = (
                            shop_variable
                            if shop_variable
                            else Path(template_file.name).stem
                        )
                        filename = f"{item_name}_C_{shop_var}.{ext}"
                        zf.writestr(filename, img_buf.getvalue())
                        count += 1

        zip_buf.seek(0)
        return zip_buf, count

    # ---- 다운로드 다이얼로그 ----
    @st.dialog("출력 설정")
    def show_save_dialog(item_files, template_files):
        st.caption("설정 후 '다운로드'를 누르면 Zip 파일이 생성됩니다.")
        fmt = "JPEG"   # 고정
        quality = 100  # 고정
        st.caption("저장 포맷: JPG(.jpg)")

        shop_variable = st.text_input(
            "Shop 구분값 (선택)",
            key="dialog_shop_var",
            help="입력 시 'Item_C_구분값.jpg' 형식으로 저장됩니다.",
        )

        if st.button("다운로드", type="primary", use_container_width=True):
            with st.spinner("이미지를 생성 중입니다..."):
                zip_buf, count = run_batch_composition(
                    item_files, template_files, fmt, quality, shop_variable
                )
            if count > 0:
                ss.download_info = {"buffer": zip_buf, "count": count}
                st.rerun()
            else:
                st.warning("생성된 이미지가 없습니다. Item이 투명 배경을 가졌는지 확인해주세요.")

    # ---- UI 레이아웃 ----
    left, right = st.columns([1, 1])

    with left:
        st.subheader("이미지 업로드")
        item_files = st.file_uploader(
            "1. Item 이미지 업로드 (누끼 딴 이미지, PNG/WEBP)",
            type=["png", "webp"],
            accept_multiple_files=True,
            key=f"item_{ss.item_uploader_key}",
        )
        if st.button("아이템 리스트 삭제"):
            ss.item_uploader_key += 1
            ss.preview_index = 0 # 인덱스 초기화
            st.rerun()

        template_files = st.file_uploader(
            "2. Template 이미지 업로드",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key=f"tpl_{ss.template_uploader_key}",
        )
        if st.button("템플릿 삭제"):
            ss.template_uploader_key += 1
            st.rerun()

    with right:
        st.subheader("이미지 설정")
        c1, c2, c3 = st.columns(3)
        c1.selectbox(
            "배치 위치",
            ["center", "top", "bottom", "left", "right",
             "top-left", "top-right", "bottom-left", "bottom-right"],
            key="anchor",
        )
        c2.selectbox(
            "리사이즈",
            [1.0, 0.9, 0.8, 0.7, 0.6],
            format_func=lambda x: f"{int(x*100)}%" if x < 1.0 else "없음",
            key="resize_ratio",
        )
        c3.selectbox("그림자 프리셋", list(SHADOW_PRESETS.keys()), key="shadow_preset")

        # 아이템 파일 목록이 바뀌면 인덱스를 0으로 리셋하여 오류 방지
        if "last_item_count" not in ss or ss.last_item_count != len(item_files):
            ss.preview_index = 0
        ss.last_item_count = len(item_files)

        # 설정 변경 시 미리보기 업데이트
        update_preview(item_files, template_files, ss.preview_index)

        # ---- 미리보기 (가장 안정적인 방법: bytes 직접 렌더) ----
        st.subheader("미리보기")
        preview_bytes = ss.get("preview_img_bytes", None)
        
        if preview_bytes and item_files:
            # 세션에 저장된 bytes를 직접 st.image에 전달합니다.
            st.image(preview_bytes, caption=f"미리보기 ({ss.preview_index + 1}/{len(item_files)})")
        else:
            # preview_bytes가 없을 때 (초기 상태 또는 생성 실패)
            st.caption("파일을 업로드하면 미리보기가 표시됩니다.")

        # ---- 미리보기 네비게이션 ----
        if item_files and len(item_files) > 1:
            col1, col2, col3 = st.columns([2, 3, 2])

            def prev_item():
                if ss.preview_index > 0:
                    ss.preview_index -= 1

            def next_item():
                if ss.preview_index < len(item_files) - 1:
                    ss.preview_index += 1
            
            with col1:
                st.button("◀ 이전", on_click=prev_item, use_container_width=True, disabled=(ss.preview_index == 0))
            
            with col2:
                 st.markdown(f"<p style='text-align: center; margin-top: 0.5rem;'>{item_files[ss.preview_index].name}</p>", unsafe_allow_html=True)

            with col3:
                st.button("다음 ▶", on_click=next_item, use_container_width=True, disabled=(ss.preview_index >= len(item_files) - 1))


        st.button(
            "생성하기",
            type="primary",
            use_container_width=True,
            disabled=(not item_files or not template_files),
            on_click=lambda: show_save_dialog(item_files, template_files),
        )

    # ---- 다운로드 버튼 ----
    if ss.get("download_info"):
        info = ss.download_info
        st.success(f"총 {info['count']}개의 이미지 생성 완료!")
        st.download_button(
            "Zip 다운로드",
            info["buffer"],
            file_name="Thumb_Craft_Results.zip",
            mime="application/zip",
            use_container_width=True,
        )
        ss.download_info = None  # 초기화


if __name__ == "__main__":
    run()


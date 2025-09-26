# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib, logging
from pathlib import Path
from typing import Optional

import streamlit as st
import gspread

# ---- 패키지 내부 모듈: 상대 임포트로 통일 ----
from .utils_common import (
    get_env, save_env_value, extract_sheet_id, sheet_link, load_env
)
from .upload_apply import collect_xlsx_files, apply_uploaded_files
from .main_controller import ShopeeAutomation

# ------------------------------------------------------------
# (선택) 버전 로깅: 디버깅 편의
# ------------------------------------------------------------
def _log_versions():
    mods = ["pandas", "openpyxl", "gspread"]
    for m in mods:
        try:
            v = importlib.import_module(m).__version__
        except Exception:
            v = "not-found"
        logging.warning(f"[VERSIONS] {m}={v}")
_log_versions()


# ------------------------------------------------------------
# 멀티 테넌트 오버라이드: 사이드바 입력 > Secrets/ENV
#  - ShopeeAutomation 내부에서 utils_common._resolve_sheet_key를 호출하므로
#    여기서 해당 함수에 '세션 오버라이드 우선' 몽키패치를 건다.
# ------------------------------------------------------------
def _install_multitenant_override():
    from . import utils_common as U  # 모듈 객체
    _orig = U._resolve_sheet_key     # 원본 보관

    def _prefer_session_override(primary_env: str, fallback_env: Optional[str] = None) -> str:
        """
        1) 사이드바 입력(세션 상태) 우선
        2) 없으면 기존 원본 로직(_resolve_sheet_key) 사용
        - primary_env/fallback_env는 보통:
          - 메인:  "GOOGLE_SHEET_KEY" / "GOOGLE_SHEETS_SPREADSHEET_ID"
          - 참조:  "REFERENCE_SHEET_KEY" / "REFERENCE_SPREADSHEET_ID"
        """
        # 세션에 담아둔 오버라이드 키/URL (빈 문자열이면 무시)
        main_raw = st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY", "").strip()
        ref_raw  = st.session_state.get("OVERRIDE_REFERENCE_SHEET_KEY", "").strip()

        def _as_key(raw: str) -> Optional[str]:
            if not raw:
                return None
            sid = extract_sheet_id(raw)  # URL/키 모두 허용
            return sid or raw

        # primary/fallback 별로 세션 오버라이드 매핑
        session_map = {
            "GOOGLE_SHEET_KEY": _as_key(main_raw),
            "GOOGLE_SHEETS_SPREADSHEET_ID": _as_key(main_raw),
            "REFERENCE_SHEET_KEY": _as_key(ref_raw),
            "REFERENCE_SPREADSHEET_ID": _as_key(ref_raw),
        }

        # 1) 세션 오버라이드가 있으면 그걸 최우선으로 사용
        if primary_env in session_map and session_map[primary_env]:
            return session_map[primary_env]
        if fallback_env in session_map and session_map[fallback_env]:
            return session_map[fallback_env]

        # 2) 없으면 기존 동작 유지 (Secrets/ENV)
        return _orig(primary_env, fallback_env)

    # 실제 패치 적용
    U._resolve_sheet_key = _prefer_session_override  # type: ignore


def run() -> None:
    """Bridge(멀티페이지) 환경에서 호출되는 진입점."""
    # (중요) 환경/설정 로드: import 시점이 아니라 실행 시점에 로드
    load_env()

    # ---- 세션 상태 초기화 ----
    defaults = {
        "upload_success": False,
        "automation_success": False,
        "download_file": None,
        # 멀티테넌트 오버라이드 기본값
        "OVERRIDE_GOOGLE_SHEET_KEY": "",
        "OVERRIDE_REFERENCE_SHEET_KEY": "",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)

    # ---- 사이드바: 멀티 테넌트 시트 오버라이드 ----
    with st.sidebar:
        st.markdown("### 🔑 Sheet Override (optional)")
        st.caption("세션(브라우저 탭)에서만 일시적으로 적용됩니다. 비워두면 Secrets/ENV 값을 사용합니다.")
        st.text_input(
            "Main Sheet Key or URL",
            key="OVERRIDE_GOOGLE_SHEET_KEY",
            placeholder="키 또는 https://docs.google.com/spreadsheets/d/... URL",
        )
        st.text_input(
            "Reference Sheet Key or URL",
            key="OVERRIDE_REFERENCE_SHEET_KEY",
            placeholder="키 또는 https://docs.google.com/spreadsheets/d/... URL",
        )
        # 세션 오버라이드가 있는지 시각 피드백
        has_main = bool(st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY"))
        has_ref  = bool(st.session_state.get("OVERRIDE_REFERENCE_SHEET_KEY"))
        st.write(
            f"Main: {'✅ Override' if has_main else '↩ Defaults'} / "
            f"Ref: {'✅ Override' if has_ref else '↩ Defaults'}"
        )

    # ---- 멀티테넌트 오버라이드 설치 (반드시 UI 구성 직후) ----
    _install_multitenant_override()

    # ---- 헤더 / 타이틀 ----
    st.title("⬆️ Copy Template")

    # ---- CSS ----
    st.markdown(
        """
<style>
html, body, [class*="st-"] { font-family: 'Inter','Noto Sans KR',sans-serif; }
div[data-testid="stAppViewContainer"] > .main .block-container {
  padding-top: 2rem; padding-bottom: 2rem; max-width: 900px;
}
.stButton>button {
  border-radius: 8px; padding: 8px 18px; font-weight: 600; border: none;
  color: white; background-color: #1A73E8; transition: background-color 0.3s ease;
}
.stButton>button:hover { background-color: #0e458c; }
.stButton>button:disabled { background-color: #E0E0E0; color: #A0A0A0; }
.stFileUploader { border: 2px dashed #E0E0E0; border-radius: 12px; padding: 20px; background-color: #F9F9F9; }
.log-container {
  background-color: #F9F9F9; border-radius: 8px; padding: 15px; margin-top: 15px;
  font-family: 'SF Mono','Menlo',monospace; font-size: 0.9em; max-height: 400px; overflow-y: auto; border: 1px solid #E0E0E0;
}
.log-success { color: #2E7D32; } .log-error { color: #C62828; } .log-warn { color: #EF6C00; } .log-info { color: #333; }
h1, h2, h3, h5 { font-weight: 700; }
.dialog-description { font-size: 0.9rem; color: #4A4A4A; margin-top: -5px; margin-bottom: 1.5rem; line-height: 1.5; }
</style>
""",
        unsafe_allow_html=True,
    )

    # ---- 설정 다이얼로그 (싱글 테넌트 기본값 저장 UI) ----
    @st.dialog("⚙️ 초기 설정")
    def settings_dialog():
        st.markdown("<h5>■ 샵 복제 시트 URL</h5>", unsafe_allow_html=True)
        st.markdown(
            """
<div class="dialog-description">
샵 복제 시트의 주소를 입력하세요.<br>
시트가 없다면 <a href="https://docs.google.com/spreadsheets/d/1l5DK-1lNGHFPfl7mbI6sTR_qU1cwHg2-tlBXzY2JhbI/edit#gid=0" target="_blank">템플릿 시트</a>에서 사본을 생성하여 입력해주세요.
</div>
""",
            unsafe_allow_html=True,
        )

        sheet_url = st.text_input(
            "Google Sheets URL",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            value=sheet_link(get_env("GOOGLE_SHEETS_SPREADSHEET_ID"))
            if get_env("GOOGLE_SHEETS_SPREADSHEET_ID")
            else "",
            label_visibility="collapsed",
        )

        st.markdown("<h5>■ 이미지 호스팅 주소</h5>", unsafe_allow_html=True)
        image_host = st.text_input(
            "Image Hosting URL",
            placeholder="예: https://dns.shopeecopy.com/",
            value=get_env("IMAGE_HOSTING_URL"),
            label_visibility="collapsed",
        )

        if st.button("저장"):
            sheet_id = extract_sheet_id(sheet_url)
            if not sheet_id:
                st.error("올바른 Google Sheets URL을 입력해주세요.")
            elif not image_host:
                st.error("이미지 호스팅 주소를 입력해주세요.")
            elif not image_host.startswith("http"):
                st.error("주소는 'http://' 또는 'https://'로 시작해야 합니다.")
            else:
                # 싱글 테넌트 기본값 저장(.env) — Cloud에선 실패할 수 있으나 로컬 편의용
                save_env_value("GOOGLE_SHEETS_SPREADSHEET_ID", sheet_id)
                save_env_value("IMAGE_HOSTING_URL", image_host)
                st.success("설정이 저장되었습니다!")
                st.rerun()

    # ---- 메인 앱 ----
    def main_application():
        col1, col2 = st.columns([0.8, 0.2])
        with col1:
            st.markdown(
                """
<p>아래 영역에 BASIC, MEDIA, SALES 엑셀 파일을 업로드하고 샵 코드를 입력한 후, 실행 버튼을 눌러주세요.</p>
""",
                unsafe_allow_html=True,
            )
        with col2:
            with st.container():
                st.write(
                    '<div style="display: flex; justify-content: flex-end; width: 100%;">',
                    unsafe_allow_html=True,
                )
                if st.button("⚙️ 설정 변경", key="edit_settings"):
                    settings_dialog()
                st.write("</div>", unsafe_allow_html=True)

        st.write("")

        # --- 입력 영역 ---
        st.subheader("1. 파일 및 샵 코드 입력")
        uploaded_files = st.file_uploader(
            "BASIC, MEDIA, SALES 파일을 한 번에 선택하거나 드래그 앤 드롭하세요.",
            type="xlsx",
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        shop_code = st.text_input(
            "샵 코드 (Shop Code) 입력",
            placeholder="예: RO, VN 등 국가 코드를 입력하세요.",
            key="shop_code_input",
        )

        is_ready = bool(uploaded_files and shop_code)

        if st.button("🚀 파일 업로드 및 전체 자동화 실행", key="run_all", disabled=not is_ready):
            # 상태 초기화
            st.session_state.upload_success = False
            st.session_state.automation_success = False
            st.session_state.download_file = None

            with st.status("자동화 실행 중...", expanded=True) as status:
                try:
                    # 1) 업로드 반영
                    st.write("1/3 - Shop SKU 파일 업로드 중...")
                    files_dict = collect_xlsx_files(uploaded_files)
                    if len(files_dict) < 3:
                        st.session_state.upload_success = False
                        status.update(label="업로드 실패", state="error", expanded=True)
                        st.error(
                            f"파일 3개(BASIC, MEDIA, SALES)를 모두 업로드해야 합니다. (현재 {len(files_dict)}개)"
                        )
                        return

                    logs = apply_uploaded_files(files_dict)
                    if any("[OK]" in log for log in logs):
                        st.session_state.upload_success = True
                        st.write("✅ 파일 업로드 완료!")
                    else:
                        status.update(label="업로드 실패", state="error", expanded=True)
                        st.error("파일을 Google Sheets에 반영하는 데 실패했습니다. 로그를 확인하세요.")
                        st.json(logs)
                        return

                    # 2) 자동화
                    st.write("2/3 - 템플릿 생성 자동화 진행 중... (Step 1~6)")
                    automation = ShopeeAutomation()
                    progress_bar = st.progress(0, text="자동화 단계를 시작합니다...")
                    log_container = st.empty()

                    success, results = automation.run_all_steps_with_progress(
                        progress_bar, log_container, shop_code
                    )
                    st.session_state.automation_success = success

                    if not success:
                        status.update(label="자동화 실패", state="error", expanded=True)
                        st.error("자동화 실행 중 오류가 발생했습니다. 위 로그를 확인하세요.")
                        return

                    # 3) 다운로드 파일 생성
                    st.write("3/3 - 최종 엑셀 파일 생성 중... (Step 7)")
                    download_data = automation.run_step7_generate_download()

                    if download_data:
                        st.session_state.download_file = download_data
                        status.update(label="🎉 모든 단계 완료!", state="complete", expanded=True)
                        st.success("모든 자동화 단계가 성공적으로 완료되었습니다!")
                    else:
                        st.session_state.automation_success = False
                        status.update(label="다운로드 파일 생성 실패", state="error", expanded=True)
                        st.error("최종 엑셀 파일을 생성하는 데 실패했습니다.")

                except Exception as e:
                    status.update(label="치명적인 오류 발생", state="error", expanded=True)
                    st.error("프로그램 실행 중 예상치 못한 심각한 오류가 발생했습니다.")
                    st.exception(e)

        st.divider()

        # --- 다운로드 섹션 ---
        st.subheader("2. 최종 파일 다운로드")
        if st.session_state.automation_success and st.session_state.download_file:
            st.download_button(
                label="⬇️ 템플릿 파일 다운로드 (.xlsx)",
                data=st.session_state.download_file,
                file_name="Shopee_Upload_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("자동화가 성공적으로 완료되면 여기에 다운로드 버튼이 나타납니다.")

    # ---- 라우팅 ----
    if not get_env("GOOGLE_SHEETS_SPREADSHEET_ID") or not get_env("IMAGE_HOSTING_URL"):
        settings_dialog()
    else:
        main_application()


# 단독 실행 지원(브릿지 없이 app.py만 직접 실행 시)
if __name__ == "__main__":
    st.set_page_config(page_title="ITEM UPLOADER", page_icon="⬆️", layout="wide")
    run()

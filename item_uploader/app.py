# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib
import logging
import os
from typing import Optional

import streamlit as st

# 내부 모듈
from .utils_common import load_env, get_env, extract_sheet_id
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
# 쿼리스트링 호환 헬퍼 (신/구 API 모두 지원)
# ------------------------------------------------------------
def set_query_params(**kwargs):
    try:
        st.query_params.update(kwargs)  # Streamlit ≥ 1.36
    except Exception:
        st.experimental_set_query_params(**kwargs)  # 구버전 백업

def get_query_params():
    try:
        return dict(st.query_params)  # Streamlit ≥ 1.36
    except Exception:
        return st.experimental_get_query_params()  # 구버전 백업


# ------------------------------------------------------------
# 멀티 테넌트 오버라이드 (메인 시트만)
#  utils_common._resolve_sheet_key에 세션 오버라이드 몽키패치
# ------------------------------------------------------------
def _install_multitenant_override():
    from . import utils_common as U
    _orig = U._resolve_sheet_key

    def _prefer_session_override(primary_env: str, fallback_env: Optional[str] = None) -> str:
        """
        세션에서 '메인 시트 키/URL'만 오버라이드.
        - Reference 시트는 오버라이드하지 않음(STRICT).
        """
        main_raw = (st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY") or "").strip()

        def _as_key(raw: str) -> Optional[str]:
            if not raw:
                return None
            sid = extract_sheet_id(raw)  # URL/키 모두 허용
            return sid or raw

        session_map = {
            "GOOGLE_SHEET_KEY": _as_key(main_raw),
            "GOOGLE_SHEETS_SPREADSHEET_ID": _as_key(main_raw),
        }

        if primary_env in session_map and session_map[primary_env]:
            return session_map[primary_env]
        if fallback_env in session_map and session_map.get(fallback_env):
            return session_map[fallback_env]

        return _orig(primary_env, fallback_env)

    U._resolve_sheet_key = _prefer_session_override  # type: ignore


def run() -> None:
    """멀티페이지(Bridge) 환경에서 호출되는 진입점."""
    load_env()

    # ── 세션 상태 초기화 ─────────────────────────────────────────
    defaults = {
        "upload_success": False,
        "automation_success": False,
        "download_file": None,
        # 메인 시트 오버라이드(키 또는 URL)
        "OVERRIDE_GOOGLE_SHEET_KEY": "",
        # 이미지 호스팅 주소(세션 우선)
        "IMAGE_HOSTING_URL_STATE": get_env("IMAGE_HOSTING_URL"),
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)

    # ── 딥링크에서 자동 복원(최초 1회 입력 목적) ────────────────────
    params = get_query_params()
    if not st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY") and params.get("main"):
        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = params["main"][0] if isinstance(params["main"], list) else params["main"]
    if not st.session_state.get("IMAGE_HOSTING_URL_STATE") and params.get("img"):
        raw_img = params["img"][0] if isinstance(params["img"], list) else params["img"]
        st.session_state["IMAGE_HOSTING_URL_STATE"] = (raw_img or "").rstrip("/")

    # 사이드바 입력 위젯 기본값은 session_state에만 세팅 (value= 사용 금지)
    st.session_state.setdefault(
        "OVERRIDE_GOOGLE_SHEET_KEY_INPUT",
        st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY", "")
    )
    st.session_state.setdefault(
        "IMAGE_HOSTING_URL_INPUT",
        st.session_state.get("IMAGE_HOSTING_URL_STATE") or get_env("IMAGE_HOSTING_URL") or ""
    )

    # ── 사이드바(항상 표시): 최소 설정 + 적용 버튼 ────────────────
    with st.sidebar:
        st.markdown("### ⚙️ 설정")
        st.markdown(
            """
            <div class="sb-help">
              샵 복제 시트의 주소를 입력하세요.<br/>
              시트가 없다면
              <a href="https://docs.google.com/spreadsheets/d/1l5DK-1lNGHFPfl7mbI6sTR_qU1cwHg2-tlBXzY2JhbI/edit#gid=0"
                 target="_blank">템플릿 시트</a>에서 사본을 생성하여 입력해주세요.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown('<div class="sb-label">샵 복제 시트 URL</div>', unsafe_allow_html=True)
        st.text_input(
            "샵 복제 시트 URL",
            key="OVERRIDE_GOOGLE_SHEET_KEY_INPUT",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            label_visibility="collapsed",
        )

        st.markdown('<div class="sb-label">이미지 호스팅 주소</div>', unsafe_allow_html=True)
        st.text_input(
            "이미지 호스팅 주소",
            key="IMAGE_HOSTING_URL_INPUT",
            placeholder="https://test.domain.com/",
            label_visibility="collapsed",
        )

        if st.button("적용", type="primary"):
            try:
                # 시트 URL/키 정규화 (비우면 오버라이드 해제 → Defaults 사용)
                raw = (st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] or "").strip()
                if raw:
                    sid = extract_sheet_id(raw)
                    if not sid:
                        raise ValueError("유효한 Google Sheets URL/키가 아닙니다.")
                    st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = sid
                else:
                    st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = ""

                # 이미지 호스팅 주소 정규화 (비우면 기본값 유지)
                host = (st.session_state["IMAGE_HOSTING_URL_INPUT"] or "").strip()
                if host:
                    if not (host.startswith("http://") or host.startswith("https://")):
                        raise ValueError("이미지 호스팅 주소는 http(s):// 로 시작해야 합니다.")
                    st.session_state["IMAGE_HOSTING_URL_STATE"] = host.rstrip("/")
                else:
                    st.session_state["IMAGE_HOSTING_URL_STATE"] = get_env("IMAGE_HOSTING_URL") or ""

                # 딥링크 저장 → 북마크/재접속 시 자동 복원 (신/구 API 호환)
                set_query_params(
                    main=st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"],
                    img=st.session_state["IMAGE_HOSTING_URL_STATE"],
                )

                st.toast("설정이 적용되었습니다 ✅")
                st.rerun()  # 최신 API
            except Exception as e:
                st.error(str(e))

    # ── 멀티테넌트 오버라이드 설치(메인만 오버라이드) ───────────────
    _install_multitenant_override()

    # ── 이미지 호스팅 주소 런타임 반영 ────────────────────────────
    # 내부 코드가 get_env('IMAGE_HOSTING_URL')로 읽으므로, os.environ에 주입
    _img_host_val = st.session_state.get("IMAGE_HOSTING_URL_STATE") or get_env("IMAGE_HOSTING_URL")
    if _img_host_val:
        os.environ["IMAGE_HOSTING_URL"] = _img_host_val

    # ── 헤더 / 타이틀 ─────────────────────────────────────────────
    st.title("⬆️ Copy Template")

    # ---- CSS (전역 + 사이드바 전용) ----
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

/* 사이드바 도움말 박스 */
[data-testid="stSidebar"] .sb-help {
  background: #F2F4F7;        /* 연한 회색 */
  color: #6B7280;             /* 텍스트 회색 */
  border: 1px solid #E5E7EB;  /* 얇은 테두리 */
  border-radius: 10px;
  padding: 10px 12px;
  line-height: 1.5;
  margin: 4px 0 14px 0;       /* 아래쪽 간격 넉넉히 */
  font-size: 0.92rem;
}
/* 라벨 느낌의 소제목 (입력창 위) */
[data-testid="stSidebar"] .sb-label {
  font-weight: 600;
  font-size: 0.95rem;
  color: #374151;
  margin: 10px 0 6px 0;       /* 라벨과 인풋 사이 간격 */
}
/* 링크 컬러 */
[data-testid="stSidebar"] .sb-help a {
  color: #2563EB;
  text-decoration: none;
}
[data-testid="stSidebar"] .sb-help a:hover {
  text-decoration: underline;
}
</style>
""",
        unsafe_allow_html=True,
    )

    # ---- 메인 앱 ----
    def main_application():
        st.markdown(
            "<p>아래 영역에 BASIC, MEDIA, SALES 엑셀 파일을 업로드하고 샵 코드를 입력한 후, 실행 버튼을 눌러주세요.</p>",
            unsafe_allow_html=True,
        )

        # --- 입력 영역 ---
        st.subheader("1. 파일 및 샵 코드 입력")
        uploaded_files = st.file_uploader(
            "BASIC, MEDIA, SALES 파일을 한 번에 선택하거나 드래그 앤 드롭하세요.",
            type="xlsx",
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        shop_code = st.text_input(
            "샵 코드 입력",
            placeholder="예: RORO, 01 등 샵 코드를 입력하세요. 커버 이미지 파일의 코드와 동일해야합니다.",
            key="shop_code_input",
        )

        is_ready = bool(uploaded_files and shop_code)

        if st.button("🚀 파일 업로드 및 실행", key="run_all", disabled=not is_ready):
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
    main_application()


# 단독 실행 지원
if __name__ == "__main__":
    st.set_page_config(page_title="ITEM UPLOADER", page_icon="⬆️", layout="wide")
    run()

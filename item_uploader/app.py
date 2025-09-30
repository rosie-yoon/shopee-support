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
    try:
        import pkg_resources  # type: ignore
        ver = pkg_resources.get_distribution("streamlit").version
    except Exception:
        ver = "unknown"
    logging.info(f"[BOOT] Streamlit={ver}")


# ------------------------------------------------------------
# URL 파라미터 헬퍼 (신/구 API 모두 지원)
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


# ------------------------------------------------------------
# 페이지 빌드
# ------------------------------------------------------------
def main():
    st.set_page_config(page_title="Copy Template", page_icon="⬆️", layout="centered")
    _log_versions()

    # ---- 환경 변수 로드 ----
    load_env()

    # ---- URL 파라미터 복원 ----
    params = get_query_params()

    # session_state 초기화: 내부 상태 키
    st.session_state.setdefault("OVERRIDE_GOOGLE_SHEET_KEY", "")
    st.session_state.setdefault("IMAGE_HOSTING_URL_STATE", get_env("IMAGE_HOSTING_URL") or "")

    # URL 파라미터(main/img) → session_state로 복원(초기 1회)
    if not st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY") and params.get("main"):
        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = (
            params["main"][0] if isinstance(params["main"], list) else params["main"]
        )
    if not st.session_state.get("IMAGE_HOSTING_URL_STATE") and params.get("img"):
        raw_img = params["img"][0] if isinstance(params["img"], list) else params["img"]
        st.session_state["IMAGE_HOSTING_URL_STATE"] = (raw_img or "").rstrip("/")

    # 사이드바 입력 위젯 기본값은 session_state에만 세팅 (value= 사용 금지)
    st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] = st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY", "")
    st.session_state["IMAGE_HOSTING_URL_INPUT"] = st.session_state.get("IMAGE_HOSTING_URL_STATE") or get_env("IMAGE_HOSTING_URL") or ""

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
            label="샵 복제 시트 URL",
            key="OVERRIDE_GOOGLE_SHEET_KEY_INPUT",
            label_visibility="collapsed",
            placeholder="https://docs.google.com/spreadsheets/d/…",
        )

        st.markdown('<div class="sb-label">이미지 호스팅 주소 (선택)</div>', unsafe_allow_html=True)
        st.text_input(
            label="이미지 호스팅 주소",
            key="IMAGE_HOSTING_URL_INPUT",
            label_visibility="collapsed",
            placeholder="https://your.cdn.host",
        )

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("적용", type="primary"):
                try:
                    # 시트 URL/키 정규화 (비우면 오버라이드 해제 → Defaults 사용)
                    raw = (st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] or "").strip()
                    if raw:
                        sid = extract_sheet_id(raw)
                        if not sid:
                            raise ValueError("유효한 Google Sheets URL/키가 아닙니다.")
                        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = sid
                        # ★ 입력창에도 정규화된 SID 반영
                        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] = sid
                    else:
                        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = ""
                        # ★ 입력창 클리어
                        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] = ""

                    # 이미지 호스팅 주소 정규화 (비우면 기본값 유지)
                    host = (st.session_state["IMAGE_HOSTING_URL_INPUT"] or "").strip()
                    if host:
                        if not (host.startswith("http://") or host.startswith("https://")):
                            raise ValueError("이미지 호스팅 주소는 http(s):// 로 시작해야 합니다.")
                        st.session_state["IMAGE_HOSTING_URL_STATE"] = host.rstrip("/")
                        # ★ 입력창에도 정규화된 호스트 반영
                        st.session_state["IMAGE_HOSTING_URL_INPUT"] = host.rstrip("/")
                    else:
                        st.session_state["IMAGE_HOSTING_URL_STATE"] = get_env("IMAGE_HOSTING_URL") or ""
                        # ★ 기본값을 입력창에도 반영
                        st.session_state["IMAGE_HOSTING_URL_INPUT"] = st.session_state["IMAGE_HOSTING_URL_STATE"]

                    # 딥링크 저장 → 북마크/재접속 시 자동 복원 (신/구 API 호환)
                    set_query_params(
                        main=st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"],
                        img=st.session_state["IMAGE_HOSTING_URL_STATE"],
                    )

                    st.toast("설정이 적용되었습니다 ✅")
                    st.rerun()  # 최신 API
                except Exception as e:
                    st.error(str(e))
        with col_b:
            if st.button("초기화"):
                st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = ""
                st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] = ""
                st.session_state["IMAGE_HOSTING_URL_STATE"] = get_env("IMAGE_HOSTING_URL") or ""
                st.session_state["IMAGE_HOSTING_URL_INPUT"] = st.session_state["IMAGE_HOSTING_URL_STATE"]
                set_query_params(main="", img=st.session_state["IMAGE_HOSTING_URL_STATE"])
                st.toast("설정이 초기화되었습니다")
                st.rerun()

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
        .log-success { color: #2E7D32; } .log-error { color: #C62828; } .log-info { color: #1565C0; }
        .sb-help { font-size: 0.9em; color: #555; margin-bottom: 6px; }
        .sb-label { font-weight: 600; margin: 12px 0 6px; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ---- 업로드 섹션 ----
    st.header("1) 파일 업로드")
    st.caption("템플릿 시트에 반영할 원본 엑셀(.xlsx)을 업로드하세요.")

    uploaded_files = st.file_uploader(
        "엑셀 파일 업로드",
        type=["xlsx"],
        accept_multiple_files=True,
        help="여러 개 업로드 가능",
    )

    if "LOGS" not in st.session_state:
        st.session_state["LOGS"] = []

    def _log(msg: str, level: str = "info"):
        st.session_state["LOGS"].append((level, msg))

    # ---- 업로드/적용 버튼 ----
    left, right = st.columns([1, 1])
    with left:
        if st.button("업로드 & 적용 실행", use_container_width=True):
            logs: list[str] = []
            try:
                files = collect_xlsx_files(uploaded_files)
                if not files:
                    st.warning("업로드된 .xlsx 파일이 없습니다.")
                # 실제 타깃 스프레드시트 확인 로그
                from .gspread_driver import open_sheet_by_env  # type: ignore
                sh = open_sheet_by_env()
                try:
                    tgt_url = getattr(sh, "url", "")
                    logs.append(f"[INFO] Target Spreadsheet: {tgt_url}")
                except Exception:
                    pass

                results = apply_uploaded_files(files, logs=logs)
                _log("파일 업로드 및 적용이 완료되었습니다.", "success")
                for ln in logs:
                    _log(ln, "info")
                if results:
                    _log(f"총 {len(results)}개 시트가 반영되었습니다.", "success")
            except Exception as e:
                _log(f"오류: {e}", "error")

    with right:
        if st.button("자동화(템플릿 생성) 실행", use_container_width=True):
            try:
                sa = ShopeeAutomation()
                sa.run()  # 내부에서 단계별 로그 출력
                _log("자동화가 완료되었습니다.", "success")
            except Exception as e:
                _log(f"자동화 오류: {e}", "error")

    # ---- 로그 출력 ----
    st.subheader("실행 로그")
    if st.session_state["LOGS"]:
        log_lines = []
        for level, msg in st.session_state["LOGS"]:
            cls = {
                "success": "log-success",
                "error": "log-error",
                "info": "log-info",
            }.get(level, "log-info")
            log_lines.append(f'<div class="{cls}">• {msg}</div>')
        st.markdown('<div class="log-container">' + "\n".join(log_lines) + "</div>", unsafe_allow_html=True)
    else:
        st.info("아직 실행 로그가 없습니다.")


if __name__ == "__main__":
    main()

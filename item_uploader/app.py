# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from typing import List

import streamlit as st

# 내부 모듈: 현재 레포에 존재하는 파일만 참조
from .utils_common import (
    load_env,
    get_env,
    extract_sheet_id,
    open_sheet_by_env,
)
from .upload_apply import collect_xlsx_files, apply_uploaded_files
from .automation_steps import (
    run_step_1,
    run_step_2,
    run_step_3,
    run_step_4,
    run_step_5,
    run_step_6,
    run_step_7,
)


# ------------------------------
# Query param helpers (신/구 API 호환)
# ------------------------------

def set_query_params(**kwargs) -> None:
    try:
        st.query_params.update(kwargs)  # Streamlit ≥ 1.36
    except Exception:
        st.experimental_set_query_params(**kwargs)  # fallback


def get_query_params() -> dict:
    try:
        return dict(st.query_params)
    except Exception:
        return st.experimental_get_query_params()


# ------------------------------
# 세션값 → ENV 주입 (메인 스프레드시트만)
# ------------------------------

def _sync_env_from_session() -> None:
    sid = (st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY") or "").strip()
    if sid:
        os.environ["GOOGLE_SHEET_KEY"] = sid
        os.environ["GOOGLE_SHEETS_SPREADSHEET_ID"] = sid
    img = (st.session_state.get("IMAGE_HOSTING_URL_STATE") or get_env("IMAGE_HOSTING_URL") or "").strip()
    if img:
        os.environ["IMAGE_HOSTING_URL"] = img


# ------------------------------
# 사이드바 설정 블록
# ------------------------------

def _sidebar_settings() -> None:
    params = get_query_params()

    # 세션 기본값
    st.session_state.setdefault("OVERRIDE_GOOGLE_SHEET_KEY", "")
    st.session_state.setdefault("IMAGE_HOSTING_URL_STATE", get_env("IMAGE_HOSTING_URL") or "")

    # URL 파라미터 복원 → 세션 (최초에만)
    if (not st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY")) and params.get("main"):
        v = params["main"][0] if isinstance(params["main"], list) else params["main"]
        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = v
    if (not st.session_state.get("IMAGE_HOSTING_URL_STATE")) and params.get("img"):
        v = params["img"][0] if isinstance(params["img"], list) else params["img"]
        st.session_state["IMAGE_HOSTING_URL_STATE"] = (v or "").rstrip("/")

    # ⚠️ 입력창: "존재하지 않을 때만" 초기화 (Enter로 인한 rerun에서 타이핑 유지)
    if "OVERRIDE_GOOGLE_SHEET_KEY_INPUT" not in st.session_state:
        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY_INPUT"] = st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY", "")
    if "IMAGE_HOSTING_URL_INPUT" not in st.session_state:
        st.session_state["IMAGE_HOSTING_URL_INPUT"] = (
            st.session_state.get("IMAGE_HOSTING_URL_STATE") or get_env("IMAGE_HOSTING_URL") or ""
        )

    with st.sidebar:
        st.markdown("### 설정")
        st.text_input(
            label="샵 복제 시트 URL",
            key="OVERRIDE_GOOGLE_SHEET_KEY_INPUT",
            placeholder="https://docs.google.com/spreadsheets/d/...",
        )
        st.text_input(
            label="이미지 호스팅 주소 (선택)",
            key="IMAGE_HOSTING_URL_INPUT",
            placeholder="https://your.cdn.host",
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("적용", type="primary"):
                try:
                    # 1) 시트 키 정규화 → 상태키만 갱신 (INPUT 키는 건드리지 않음)
                    raw = (st.session_state.get("OVERRIDE_GOOGLE_SHEET_KEY_INPUT") or "").strip()
                    if raw:
                        sid = extract_sheet_id(raw)
                        if not sid:
                            raise ValueError("유효한 Google Sheets URL/키가 아닙니다.")
                        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = sid
                    else:
                        st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = ""

                    # 2) 이미지 호스트 정규화 → 상태키만 갱신
                    host = (st.session_state.get("IMAGE_HOSTING_URL_INPUT") or "").strip()
                    if host:
                        if not (host.startswith("http://") or host.startswith("https://")):
                            raise ValueError("이미지 호스팅 주소는 http(s):// 로 시작해야 합니다.")
                        st.session_state["IMAGE_HOSTING_URL_STATE"] = host.rstrip("/")
                    else:
                        st.session_state["IMAGE_HOSTING_URL_STATE"] = get_env("IMAGE_HOSTING_URL") or ""

                    # 3) 딥링크 갱신
                    set_query_params(
                        main=st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"],
                        img=st.session_state["IMAGE_HOSTING_URL_STATE"],
                    )

                    # 4) 다음 렌더에서 입력창을 최신 상태로 재초기화하기 위해 INPUT 키 제거 후 rerun
                    st.session_state.pop("OVERRIDE_GOOGLE_SHEET_KEY_INPUT", None)
                    st.session_state.pop("IMAGE_HOSTING_URL_INPUT", None)

                    st.toast("설정이 적용되었습니다 ✅")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
        with c2:
            if st.button("초기화"):
                st.session_state["OVERRIDE_GOOGLE_SHEET_KEY"] = ""
                st.session_state["IMAGE_HOSTING_URL_STATE"] = get_env("IMAGE_HOSTING_URL") or ""

                set_query_params(main="", img=st.session_state["IMAGE_HOSTING_URL_STATE"])

                # 입력창을 초기 표시값으로 되돌리기 위해 키 제거
                st.session_state.pop("OVERRIDE_GOOGLE_SHEET_KEY_INPUT", None)
                st.session_state.pop("IMAGE_HOSTING_URL_INPUT", None)

                st.toast("설정이 초기화되었습니다")
                st.rerun()


# ------------------------------
# 자동화 실행 래퍼
# ------------------------------

def _run_automation(logs: List[str]) -> None:
    try:
        logs.append("[STEP] 1: 준비/검증 시작")
        run_step_1()
        logs.append("[STEP] 2: 원본 정리")
        run_step_2()
        logs.append("[STEP] 3: 매핑/병합")
        run_step_3()
        logs.append("[STEP] 4: 이미지 정리")
        run_step_4()
        logs.append("[STEP] 5: 템플릿 구성")
        run_step_5()
        logs.append("[STEP] 6: 산출물 생성")
        run_step_6()
        logs.append("[STEP] 7: 최종 템플릿 내보내기")
        out = run_step_7()
        if out:
            logs.append("[OK] 최종 템플릿 파일이 생성되었습니다.")
        else:
            logs.append("[WARN] 최종 템플릿 파일이 비어있습니다.")
    except Exception as e:
        logs.append(f"[ERROR] 자동화 실패: {e}")
        raise


# ------------------------------
# 메인 렌더
# ------------------------------

def _render() -> None:
    # 멀티페이지 환경에서 상위 스크립트가 이미 set_page_config를 호출했을 수 있음 → 호출하지 않음
    load_env()
    _sidebar_settings()
    _sync_env_from_session()

    st.title("Copy Template")

    st.header("1) 파일 업로드")
    st.caption("템플릿 시트에 반영할 원본 엑셀(.xlsx)을 업로드하세요.")

    uploaded_files = st.file_uploader(
        "엑셀 파일 업로드",
        type=["xlsx"],
        accept_multiple_files=True,
        help="여러 개 업로드 가능",
    )

    # 로그 버퍼
    if "LOGS" not in st.session_state:
        st.session_state["LOGS"] = []

    def _log(msg: str, level: str = "info") -> None:
        st.session_state["LOGS"].append((level, msg))

    col1, col2 = st.columns(2)
    with col1:
        if st.button("업로드 & 적용 실행", use_container_width=True):
            logs: List[str] = []
            try:
                files = collect_xlsx_files(uploaded_files)
                if not files:
                    st.warning("업로드된 .xlsx 파일이 없습니다.")
                # 대상 스프레드시트 로그
                try:
                    sh = open_sheet_by_env()
                    tgt_url = getattr(sh, "url", "")
                    logs.append(f"[INFO] Target Spreadsheet: {tgt_url}")
                except Exception:
                    pass

                results = apply_uploaded_files(files, logs=logs)
                for ln in logs:
                    _log(ln, "info")
                _log("파일 업로드 및 적용이 완료되었습니다.", "success")
                if results:
                    _log(f"총 {len(results)}개 시트가 반영되었습니다.", "success")
            except Exception as e:
                _log(f"오류: {e}", "error")

    with col2:
        if st.button("자동화(템플릿 생성) 실행", use_container_width=True):
            try:
                logs: List[str] = []
                _run_automation(logs)
                for ln in logs:
                    _log(ln, "info")
                _log("자동화가 완료되었습니다.", "success")
            except Exception as e:
                _log(f"자동화 오류: {e}", "error")

    # 로그 출력
    st.subheader("실행 로그")
    if st.session_state["LOGS"]:
        lines: List[str] = []
        for level, msg in st.session_state["LOGS"]:
            color = {"success": "#2E7D32", "error": "#C62828", "info": "#1565C0"}.get(level, "#1565C0")
            lines.append(f'<div style="color:{color}">• {msg}</div>')
        html = (
            '<div style="background:#F9F9F9;border:1px solid #eee;border-radius:8px;padding:12px;max-height:360px;overflow:auto">'
            + "\n".join(lines)
            + "</div>"
        )
        st.markdown(html, unsafe_allow_html=True)
    else:
        st.info("아직 실행 로그가 없습니다.")


# ------------------------------
# 공개 엔트리포인트 (pages/*에서 import)
# ------------------------------

def run() -> None:
    _render()


# 로컬 실행 지원
if __name__ == "__main__":
    run()

# -*- coding: utf-8 -*-
"""
utils_common.py (CLEAN)
- Cloud/Local 환경에서 안정적으로 Google Sheets 인증/접근하도록 정리한 버전
- 우선순위: Streamlit Secrets(service account) → ENV(JSON 문자열) → 로컬 OAuth(client_secret.json)
- 기존 유틸/인터페이스 최대한 유지
"""
from __future__ import annotations

import os
import re
import time
import json
from pathlib import Path
from typing import Optional, Dict, Callable, List

import gspread
from gspread.exceptions import WorksheetNotFound
from dotenv import load_dotenv

# -----------------------------
# 환경변수 로딩
# -----------------------------
def load_env():
    """여러 위치에서 .env 탐색하여 로드"""
    base = Path(__file__).resolve().parent
    for p in [base / ".env", base.parent / ".env", Path.cwd() / ".env"]:
        if p.exists():
            load_dotenv(p, override=True)
            return
    load_dotenv(override=True)  # fallback: 시스템 환경변수만


def get_env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def get_bool_env(name: str, default: bool = False) -> bool:
    val = os.getenv(name, "").strip().lower()
    if val in ["1", "true", "yes", "y"]:
        return True
    if val in ["0", "false", "no", "n"]:
        return False
    return default


def save_env_value(name: str, value: str, search_paths: Optional[List[Path]] = None) -> bool:
    """
    .env에 name=value를 저장(있으면 교체, 없으면 추가).
    - Cloud(읽기전용)에서는 실패할 수 있으므로 False 반환 가능.
    - 로컬 개발 편의용 유틸. 반환값으로 성공 여부만 알려줌.
    """
    name = str(name).strip()
    value = str(value)
    if not name:
        return False

    # 탐색 경로: 현재 파일 근처 → 상위 → CWD
    base = Path(__file__).resolve().parent
    candidates = search_paths or [base / ".env", base.parent / ".env", Path.cwd() / ".env"]

    env_path: Optional[Path] = None
    for p in candidates:
        try:
            if p.exists() and p.is_file():
                env_path = p
                break
        except Exception:
            continue
    if env_path is None:
        # 첫 후보에 새로 생성 시도
        env_path = candidates[0]

    try:
        lines: List[str] = []
        if env_path.exists():
            lines = env_path.read_text(encoding="utf-8").splitlines()

        pattern = re.compile(rf"^\s*{re.escape(name)}\s*=\s*.*$")
        replaced = False
        for i, line in enumerate(lines):
            if pattern.match(line):
                lines[i] = f"{name}={value}"
                replaced = True
                break
        if not replaced:
            lines.append(f"{name}={value}")

        env_path.parent.mkdir(parents=True, exist_ok=True)
        env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return True
    except Exception:
        # 쓰기 불가(예: 클라우드 읽기전용 등)
        return False


# -----------------------------
# gspread 인증/시트 접근
# -----------------------------
def with_retry(fn: Callable, retries: int = 3, delay: float = 2.0):
    """API 요청 재시도 래퍼"""
    last_err = None
    for _ in range(retries):
        try:
            return fn()
        except Exception as e:
            last_err = e
            time.sleep(delay)
    if last_err:
        raise last_err


def _service_account_from_streamlit_or_env() -> Optional[gspread.Client]:
    """Streamlit Secrets 또는 ENV의 서비스계정 JSON으로 gspread 클라이언트를 생성.
    둘 다 없으면 None 반환.
    """
    # 1) Streamlit Secrets
    try:
        import streamlit as st  # type: ignore
        if "gcp_service_account" in st.secrets:
            creds_info = st.secrets["gcp_service_account"]  # dict
            return gspread.service_account_from_dict(dict(creds_info))
    except Exception:
        pass

    # 2) ENV JSON 문자열
    try:
        env_json = os.getenv("GCP_SERVICE_ACCOUNT_JSON", "").strip()
        if env_json:
            creds_info = json.loads(env_json)
            return gspread.service_account_from_dict(creds_info)
    except Exception:
        pass

    return None


def _resolve_sheet_key(primary_env: str, fallback_env: Optional[str] = None) -> str:
    """시트 키를 secrets/ENV에서 해석. URL/키 모두 허용."""
    # Streamlit secrets → ENV 순서로 조회
    val: Optional[str] = None
    try:
        import streamlit as st  # type: ignore
        # secrets 우선 (예: GOOGLE_SHEET_KEY)
        if primary_env in st.secrets:
            val = str(st.secrets.get(primary_env, "") or "").strip()
        elif fallback_env and fallback_env in st.secrets:
            val = str(st.secrets.get(fallback_env, "") or "").strip()
    except Exception:
        pass

    if not val:
        val = os.getenv(primary_env, "").strip()
    if not val and fallback_env:
        val = os.getenv(fallback_env, "").strip()

    if not val:
        raise RuntimeError(f"{primary_env} (또는 {fallback_env}) 가 secrets/ENV에 설정되어 있지 않습니다.")

    # URL or raw key 지원
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9\-_]+)", val)
    return m.group(1) if m else val


def open_sheet_by_env():
    """
    실행 환경에 따라 인증 방식을 선택하여 메인 스프레드시트를 연다.
    - Cloud: Streamlit Secrets(service account) → ENV(JSON)
    - Local: 위가 없으면 OAuth(client_secret.json/token.json)
    키 이름 호환: GOOGLE_SHEET_KEY → (fallback) GOOGLE_SHEETS_SPREADSHEET_ID
    """
    load_env()

    # 1) 인증: 서비스계정 우선
    gc = _service_account_from_streamlit_or_env()

    # 2) 로컬 OAuth 폴백
    if gc is None:
        here = Path(__file__).resolve().parent
        cred_path = here / "client_secret.json"
        token_path = here / "token.json"
        gc = gspread.oauth(
            credentials_filename=str(cred_path),
            authorized_user_filename=str(token_path),
        )

    # 3) 시트 키 해석 (URL/키 모두 허용)
    sheet_key = _resolve_sheet_key(
        primary_env="GOOGLE_SHEET_KEY",
        fallback_env="GOOGLE_SHEETS_SPREADSHEET_ID",
    )
    return gc.open_by_key(sheet_key)


def open_ref_by_env():
    """레퍼런스 시트(선택)를 연다. 없으면 None 반환.
    키 이름 호환: REFERENCE_SHEET_KEY → (fallback) REFERENCE_SPREADSHEET_ID
    """
    load_env()

    # 인증 재사용
    gc = _service_account_from_streamlit_or_env()
    if gc is None:
        here = Path(__file__).resolve().parent
        cred_path = here / "client_secret.json"
        token_path = here / "token.json"
        try:
            gc = gspread.oauth(
                credentials_filename=str(cred_path),
                authorized_user_filename=str(token_path),
            )
        except Exception:
            # 레퍼런스 시트는 optional. 인증 실패 시 None 처리
            return None

    # 키가 없으면 optional로 간주하고 None
    try:
        ref_key = _resolve_sheet_key(
            primary_env="REFERENCE_SHEET_KEY",
            fallback_env="REFERENCE_SPREADSHEET_ID",
        )
    except Exception:
        return None

    try:
        return gc.open_by_key(ref_key)
    except Exception:
        return None


def safe_worksheet(sh, name: str):
    if not sh:
        raise ValueError(f"Spreadsheet object is not valid. Cannot get worksheet '{name}'.")
    try:
        return with_retry(lambda: sh.worksheet(name))
    except WorksheetNotFound:
        raise


# -----------------------------
# 문자열/헤더 정규화
# -----------------------------
def norm(s: str) -> str:
    return (
        str(s or "")
        .strip()
        .lower()
        .replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
    )


def header_key(s: str) -> str:
    """헤더 비교용: 영숫자+하이픈만 남김"""
    return re.sub(r"[^a-z0-9\-]+", "", norm(s))


def hex_to_rgb01(hex_str: str) -> Dict[str, float]:
    """#RRGGBB → {red,green,blue} (0~1 float)"""
    hex_str = hex_str.lstrip("#")
    if len(hex_str) != 6:
        return {"red": 1, "green": 1, "blue": 0.7}
    r, g, b = tuple(int(hex_str[i : i + 2], 16) for i in (0, 2, 4))
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}


def extract_sheet_id(s: str) -> str | None:
    s = (s or "").strip()
    if re.fullmatch(r"[A-Za-z0-9\-_]{25,}", s):
        return s
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9\-_]+)", s)
    if m:
        return m.group(1)
    return None


def sheet_link(sid: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sid}/edit"


# -----------------------------
# 카테고리 관련 유틸
# -----------------------------
def strip_category_id(cat: str) -> str:
    """ '101814 - Home & Living/...' -> 'Home & Living/...' """
    s = str(cat or "")
    m = re.match(r"^\s*\d+\s*-\s*(.+)$", s)
    return m.group(1) if m else s


def top_of_category(cat: str) -> Optional[str]:
    """ TopLevel 추출 """
    if not cat:
        return None
    tail = strip_category_id(cat)
    for sep in ["/", ">", "|", "\\"]:
        if sep in tail:
            tail = tail.split(sep, 1)[0]
            break
    tail = tail.strip()
    return tail.lower() if tail else None


# -----------------------------
# TEM 시트명
# -----------------------------
def get_tem_sheet_name() -> str:
    return get_env("TEM_OUTPUT_SHEET_NAME", "TEM_OUTPUT")

# -*- coding: utf-8 -*-
"""
utils_common.py (FINAL, STRICT)
- Streamlit Cloud/Local 에서 안정적인 Google Sheets 접근
- 인증: Service Account만 허용 (Secrets[gcp_service_account] 또는 ENV[GCP_SERVICE_ACCOUNT_JSON])
- 시트 키 해석: secrets/env 값이 URL/키 어느 쪽이든 허용
- with_retry: 429/5xx에 지수 백오프 + 지터
"""
from __future__ import annotations

import os
import re
import json
import time
import random
from pathlib import Path
from typing import Optional, Dict, Callable, List

import gspread
from gspread.exceptions import WorksheetNotFound
from dotenv import load_dotenv


# -----------------------------
# ENV / Secrets
# -----------------------------
def load_env() -> None:
    """여러 위치에서 .env 탐색하여 로드"""
    base = Path(__file__).resolve().parent
    for p in [base / ".env", base.parent / ".env", Path.cwd() / ".env"]:
        if p.exists():
            load_dotenv(p, override=True)
            return
    load_dotenv(override=True)  # fallback


def get_env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def get_bool_env(name: str, default: bool = False) -> bool:
    v = os.getenv(name, "").strip().lower()
    if v in ("1", "true", "yes", "y"):
        return True
    if v in ("0", "false", "no", "n"):
        return False
    return default


def save_env_value(name: str, value: str, search_paths: Optional[List[Path]] = None) -> bool:
    """
    .env에 name=value 저장(있으면 교체, 없으면 추가).
    - Cloud(읽기전용)에서는 실패할 수 있으므로 False 반환 가능.
    - 로컬 개발 편의 유틸.
    """
    name = str(name).strip()
    value = str(value)
    if not name:
        return False

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
        env_path = candidates[0]

    try:
        lines: List[str] = []
        if env_path.exists():
            lines = env_path.read_text(encoding="utf-8").splitlines()

        patt = re.compile(rf"^\s*{re.escape(name)}\s*=\s*.*$")
        replaced = False
        for i, line in enumerate(lines):
            if patt.match(line):
                lines[i] = f"{name}={value}"
                replaced = True
                break
        if not replaced:
            lines.append(f"{name}={value}")

        env_path.parent.mkdir(parents=True, exist_ok=True)
        env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return True
    except Exception:
        return False


# -----------------------------
# Retry (429/5xx 지수 백오프)
# -----------------------------
def with_retry(
    fn: Callable,
    retries: int = 8,
    base_delay: float = 2.0,
    backoff: float = 1.8,
    max_delay: float = 65.0,
):
    """
    gspread 호출용 재시도 래퍼
    - 429(quota/rate) 또는 500/502/503/504에서 지수 백오프 + 지터로 재시도
    - 그 외 에러는 즉시 전파
    """
    last_err = None
    delay = base_delay
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as e:
            code = getattr(getattr(e, "response", None), "status_code", None)
            msg = (str(e) or "").lower()
            is_rate = (code in (429, 500, 502, 503, 504)) or ("quota" in msg) or ("rate" in msg)

            if is_rate and attempt < retries:
                sleep_for = min(max_delay, delay + random.uniform(0, delay * 0.3))
                time.sleep(sleep_for)
                delay = min(max_delay, delay * backoff)
                last_err = e
                continue

            last_err = e
            break
    raise last_err


# -----------------------------
# gspread 인증 / 시트 키 해석
# -----------------------------
def _service_account_from_streamlit_or_env() -> Optional[gspread.Client]:
    """
    Streamlit Secrets 또는 ENV(GCP_SERVICE_ACCOUNT_JSON)의 서비스계정 JSON으로 Client 생성.
    - 둘 다 없으면 None
    """
    # 1) Streamlit Secrets
    try:
        import streamlit as st  # type: ignore
        if "gcp_service_account" in st.secrets:
            creds_info = dict(st.secrets["gcp_service_account"])
            return gspread.service_account_from_dict(creds_info)
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
    """
    시트 키를 secrets/ENV에서 해석. URL/키 모두 허용.
    - 우선순위: Streamlit Secrets → ENV
    """
    val: Optional[str] = None
    try:
        import streamlit as st  # type: ignore
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

    m = re.search(r"/spreadsheets/d/([A-Za-z0-9\-_]+)", val)
    return m.group(1) if m else val


def open_sheet_by_env():
    """
    메인 스프레드시트 오픈 (STRICT)
    - Service Account만 허용 (Secrets[gcp_service_account] / ENV[GCP_SERVICE_ACCOUNT_JSON])
    - 폴백(OAuth client_secret.json) 없음
    """
    load_env()
    gc = _service_account_from_streamlit_or_env()
    if gc is None:
        raise RuntimeError(
            "[AUTH] 서비스계정 인증 정보를 찾지 못했습니다. "
            "Streamlit Secrets에 [gcp_service_account] 블록을 추가하거나 "
            "환경변수 GCP_SERVICE_ACCOUNT_JSON을 설정하세요."
        )
    sheet_key = _resolve_sheet_key("GOOGLE_SHEET_KEY", "GOOGLE_SHEETS_SPREADSHEET_ID")
    return gc.open_by_key(sheet_key)


def open_ref_by_env():
    """
    참조 시트(REFERENCE_SHEET_KEY)만 연다 (STRICT)
    - 실패 시 명확한 예외 (None/폴백 금지)
    """
    load_env()
    gc = _service_account_from_streamlit_or_env()
    if gc is None:
        raise RuntimeError(
            "[AUTH] 서비스계정 인증 정보를 찾지 못했습니다. "
            "Streamlit Secrets 또는 GCP_SERVICE_ACCOUNT_JSON을 확인하세요."
        )
    ref_key = _resolve_sheet_key("REFERENCE_SHEET_KEY", "REFERENCE_SPREADSHEET_ID")
    try:
        return gc.open_by_key(ref_key)
    except Exception as e:
        raise RuntimeError(
            f"[REF] REFERENCE_SHEET_KEY로 참조 시트를 열지 못했습니다 (key={ref_key}). "
            f"→ 키 값/서비스계정 공유 권한을 확인하세요. 원인: {e}"
        ) from e


def safe_worksheet(sh, name: str):
    if not sh:
        raise ValueError(f"Spreadsheet object is not valid. Cannot get worksheet '{name}'.")
    try:
        return with_retry(lambda: sh.worksheet(name))
    except WorksheetNotFound:
        raise


def get_or_create_worksheet(sh, name: str, rows: int = 100, cols: int = 26):
    """워크시트가 없으면 생성하여 반환"""
    if not sh:
        raise ValueError(f"Spreadsheet object is not valid. Cannot get or create worksheet '{name}'.")
    try:
        return with_retry(lambda: sh.worksheet(name))
    except WorksheetNotFound:
        return with_retry(lambda: sh.add_worksheet(title=name, rows=rows, cols=cols))


# -----------------------------
# 문자열/헤더 정규화 & 유틸
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
# 카테고리 관련
# -----------------------------
def strip_category_id(cat: str) -> str:
    """ '101814 - Home & Living/...' -> 'Home & Living/...' """
    s = str(cat or "")
    m = re.match(r"^\s*\d+\s*-\s*(.+)$", s)
    return m.group(1) if m else s


def top_of_category(cat: str) -> Optional[str]:
    """ TopLevel 추출 (구분자: '/', '>', '|', '\\') """
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

# -*- coding: utf-8 -*-
"""
utils_common.py
- TEM Uploader 프로젝트 공통 유틸 함수 모음
- Step1~Step5, Streamlit UI에서 공유
"""

from __future__ import annotations
import os
import re
import time
import json
from pathlib import Path
from typing import Optional, List, Dict, Callable

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

def get_bool_env(name: str, default: bool=False) -> bool:
    val = os.getenv(name, "").strip().lower()
    if val in ["1","true","yes","y"]: return True
    if val in ["0","false","no","n"]: return False
    return default

def _env_path() -> str:
    """ .env 파일 경로 추정 (프로젝트 루트 우선) """
    here = Path(__file__).resolve().parent
    p1 = here / ".env"
    if p1.exists():
        return str(p1)
    return str(Path.cwd() / ".env")

def save_env_value(key: str, value: str):
    """ 단순 .env 업데이트: 키 있으면 교체, 없으면 추가 """
    path = _env_path()
    kv = {}
    if Path(path).exists():
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.rstrip("\n")
                if not line or line.strip().startswith("#"): continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    kv[k.strip()] = v.strip()
    kv[key] = value
    lines = [f"{k}={v}" for k, v in kv.items()]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

# -----------------------------
# gspread 인증/시트 접근
# -----------------------------
def with_retry(fn: Callable, retries: int=3, delay: float=2.0):
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

def open_sheet_by_env():
    """
    (수정) 시트 열기 직전에 항상 load_env()를 호출하여 최신 설정을 반영합니다.
    """
    load_env()  # ★★★ 이 부분이 핵심 수정 사항입니다.
    ss_id = get_env("GOOGLE_SHEETS_SPREADSHEET_ID")
    if not ss_id:
        raise RuntimeError("GOOGLE_SHEETS_SPREADSHEET_ID not set in .env")
    
    # 인증 파일 경로를 더 안정적으로 찾도록 수정
    cred_path = Path(__file__).resolve().parent / "client_secret.json"
    token_path = Path(__file__).resolve().parent / "token.json"
    
    gc = gspread.oauth(
        credentials_filename=str(cred_path),
        authorized_user_filename=str(token_path)
    )
    return gc.open_by_key(ss_id)

def open_ref_by_env():
    """
    (수정) 레퍼런스 시트 열기 직전에도 항상 load_env()를 호출합니다.
    """
    load_env()  # ★★★ 이 부분이 핵심 수정 사항입니다.
    ref_id = get_env("REFERENCE_SPREADSHEET_ID")
    if not ref_id:
        # 레퍼런스 시트는 없을 수도 있으므로 오류 대신 None을 반환하도록 처리
        return None
        
    cred_path = Path(__file__).resolve().parent / "client_secret.json"
    token_path = Path(__file__).resolve().parent / "token.json"

    gc = gspread.oauth(
        credentials_filename=str(cred_path),
        authorized_user_filename=str(token_path)
    )
    return gc.open_by_key(ref_id)


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

def hex_to_rgb01(hex_str: str) -> Dict[str,float]:
    """#RRGGBB → {red,green,blue} (0~1 float)"""
    hex_str = hex_str.lstrip("#")
    if len(hex_str) != 6: return {"red":1,"green":1,"blue":0.7}
    r,g,b = tuple(int(hex_str[i:i+2],16) for i in (0,2,4))
    return {"red":r/255.0,"green":g/255.0,"blue":b/255.0}

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
    if not cat: return None
    tail = strip_category_id(cat)
    for sep in ["/", ">", "|", "\\"]:
        if sep in tail:
            tail = tail.split(sep,1)[0]
            break
    tail = tail.strip()
    return tail.lower() if tail else None

# -----------------------------
# TEM 시트명
# -----------------------------
def get_tem_sheet_name() -> str:
    return get_env("TEM_OUTPUT_SHEET_NAME", "TEM_OUTPUT")


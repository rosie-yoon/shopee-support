# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
from io import BytesIO
from typing import List, Optional
from zipfile import ZipFile, ZIP_DEFLATED

import gspread
from gspread.utils import rowcol_to_a1
import pandas as pd
from openpyxl import load_workbook

# 프로젝트 공통 유틸
from .utils_common import open_sheet_by_env, safe_worksheet, with_retry

# ======================================================
# XLSX 파싱 안정화 (Shopee/엑셀 시트뷰 불일치 대응 + 숨김행/메타 안전 처리)
# ======================================================

def _sanitize_xlsx_for_openpyxl(file_bytes: bytes) -> BytesIO:
    """
    - 일부 XLSX에서 openpyxl이 sheetViews/pane의 enum 검증에서 실패하는 문제를 예방.
    - 모든 worksheet XML의 <sheetViews ...> 블록 및 <pane/> 단독 태그를 제거.
    - 네임스페이스 접두사와 속성이 붙은 경우까지 전부 제거하도록 정규식 강화.
    """
    sanitized_buffer = BytesIO()
    try:
        with ZipFile(BytesIO(file_bytes), 'r') as zin, ZipFile(sanitized_buffer, 'w', ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                buffer = zin.read(item.filename)
                if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                    # 안전한 디코드 (깨진 문자 무시)
                    xml_content = buffer.decode('utf-8', errors='ignore')

                    # 1) <sheetViews ...> ... </sheetViews> 제거 (네임스페이스/속성 포함)
                    cleaned_xml = re.sub(
                        r'<(?:\w+:)?sheetViews\b[^>]*>.*?</(?:\w+:)?sheetViews>',
                        '',
                        xml_content,
                        flags=re.DOTALL | re.IGNORECASE,
                    )

                    # 2) <pane .../> 단독 태그 추가 제거 (방역)
                    cleaned_xml = re.sub(
                        r'<(?:\w+:)?pane\b[^>]*/\s*>',
                        '',
                        cleaned_xml,
                        flags=re.DOTALL | re.IGNORECASE,
                    )

                    buffer = cleaned_xml.encode('utf-8')
                # 원본/수정된 파일 모두 새 ZIP에 기록
                zout.writestr(item, buffer)
        sanitized_buffer.seek(0)
        return sanitized_buffer
    except Exception:
        # 문제가 생기면 원본 반환 (폴백)
        return BytesIO(file_bytes)


# ------------------------
# robust openpyxl / pandas
# ------------------------

def _count_nonempty_cells(ws, max_rows: int = 500, max_cols: int = 200) -> int:
    """시트 앞쪽 일부에서 non-empty 셀 수를 세어 데이터가 가장 많은 시트를 고른다."""
    rows = min(ws.max_row or 0, max_rows)
    cols = min(ws.max_column or 0, max_cols)
    cnt = 0
    for r in ws.iter_rows(min_row=1, max_row=rows, min_col=1, max_col=cols, values_only=True):
        for c in r:
            if c not in (None, '', ' '):
                cnt += 1
    return cnt


def _read_with_openpyxl(sanitized_bio: BytesIO, debug: bool = False) -> List[List[str]]:
    """정리된 바이트로 openpyxl을 이용해 값을 읽는다.
    - 숨김 행도 포함해서 읽고, 나중에 완전 빈 행만 제거
    - 데이터가 가장 많은 시트를 선택
    """
    data: List[List[str]] = []
    try:
        sanitized_bio.seek(0)
        wb = load_workbook(sanitized_bio, data_only=True, read_only=True)
        ws = max(wb.worksheets, key=lambda s: _count_nonempty_cells(s))
        if debug:
            print(f"[DEBUG] openpyxl target sheet = {ws.title}")

        for row in ws.iter_rows(values_only=True):
            str_row = [str(cell) if cell is not None else "" for cell in row]
            data.append(str_row)

        # 완전 빈 행 제거
        data = [r for r in data if any(v.strip() for v in r)]
    except Exception as e:
        if debug:
            print(f"[DEBUG] openpyxl read failed → {e}")
        data = []
    return data


def _read_with_pandas_fallback(sanitized_bio: BytesIO, debug: bool = False) -> List[List[str]]:
    """pandas 폴백을 정리된 바이트로 수행. calamine가 있으면 우선 사용."""
    # 1) calamine 엔진 시도 (설치되어 있으면 openpyxl 의존 회피)
    try:
        sanitized_bio.seek(0)
        df = pd.read_excel(sanitized_bio, header=None, dtype=str, engine="calamine").fillna('')  # type: ignore[arg-type]
        if debug:
            print("[DEBUG] pandas calamine engine used")
        return df.values.tolist()
    except Exception as e:
        if debug:
            print(f"[DEBUG] pandas calamine not used → {e}")

    # 2) 기본 엔진 (보통 openpyxl)로 재시도 — 이미 sanitize됐으므로 상대적으로 안전
    try:
        sanitized_bio.seek(0)
        df = pd.read_excel(sanitized_bio, header=None, dtype=str).fillna('')
        if debug:
            print("[DEBUG] pandas default engine used")
        return df.values.tolist()
    except Exception as e:
        if debug:
            print(f"[DEBUG] pandas default failed → {e}")
        return []


def read_xlsx_values(bio: BytesIO, debug: bool = True) -> List[List[str]]:
    """
    업로드된 XLSX BytesIO를 안정적으로 파싱하여 2D 리스트로 반환합니다.
    1) XML sanitize로 sheetViews/pane 제거
    2) openpyxl로 시도(숨김행 포함/실데이터 많은 시트 선택) → 비정상 시 pandas로 폴백
    3) Shopee 메타 데이터 행 최소 제거(최대 1회씩)
    """
    original_bytes = bio.getvalue()
    sanitized_bio = _sanitize_xlsx_for_openpyxl(original_bytes)

    # 1) openpyxl 경로
    data = _read_with_openpyxl(sanitized_bio, debug=debug)

    # 2) 결과가 비정상적이면 pandas 폴백 (정리된 바이트 기반)
    is_data_invalid = (not data) or (len(data) == 1 and len(data[0]) <= 1)
    if is_data_invalid:
        data = _read_with_pandas_fallback(sanitized_bio, debug=debug)
        if not data:
            if debug:
                print("[DEBUG] both readers failed → return []")
            return []

    # 3) Shopee 메타 데이터 행 최소 제거 (과제거 방지)
    # et_title_* 헤더가 맨 윗줄에 있으면 한 번만 제거
    if data and data[0] and any(str(c).startswith('et_title_') for c in data[0]):
        data.pop(0)
    # 구역 라벨(basic_info/media_info/sales_info)이 맨 윗줄에 있으면 한 번만 제거
    if data and data[0] and str(data[0][0]).strip() in ('basic_info', 'media_info', 'sales_info'):
        data.pop(0)

    if debug:
        print(f"[DEBUG] final rows={len(data)} preview={data[:3]}")

    return data


# ------------------------------------------------------
# Google Sheet 쓰기 (기존 로직 유지)
# ------------------------------------------------------

def _write_values_to_sheet(sh: gspread.Spreadsheet, tab: str, values: List[List], logs: List[str]) -> None:
    rows = len(values)
    cols = max((len(r) for r in values), default=0)
    logs.append(f"[INFO] {tab}: parsed shape = {rows}x{cols}")

    if rows == 0 or cols == 0:
        logs.append(f"[WARN] {tab}: 입력 데이터가 비어 있어 skip")
        return

    try:
        ws = safe_worksheet(sh, tab)
        with_retry(lambda: ws.clear())
    except Exception:
        ws = with_retry(lambda: sh.add_worksheet(title=tab, rows=max(rows + 10, 100), cols=max(cols + 5, 26)))

    if ws.row_count < rows or ws.col_count < cols:
        with_retry(lambda: ws.resize(rows=rows + 10, cols=cols + 5))

    chunk_rows = int(os.getenv("UPLOAD_CHUNK_ROWS", "0") or "0")
    if chunk_rows and chunk_rows > 0:
        for i in range(0, rows, chunk_rows):
            chunk = values[i:i + chunk_rows]
            start_a1 = rowcol_to_a1(i + 1, 1)
            with_retry(lambda: ws.update(start_a1, chunk, raw=True))
    else:
        end_a1 = rowcol_to_a1(rows, cols)
        with_retry(lambda: ws.update(f"A1:{end_a1}", values, raw=True))

    logs.append(f"[OK] {tab}: {rows}x{cols} 적용 완료")


# ------------------------------------------------------
# 파일명 → 탭 자동 라우팅 (기존 로직 유지)
# ------------------------------------------------------

def _target_tab(filename: str) -> Optional[str]:
    low = filename.lower()
    if "basic" in low:
        return "BASIC"
    if "media" in low:
        return "MEDIA"
    if "sales" in low:
        return "SALES"
    return None


# ------------------------------------------------------
# 업로드 반영 엔트리 (기존 로직 유지 / 로그 보강)
# ------------------------------------------------------

def apply_uploaded_files(files: dict[str, BytesIO]) -> list[str]:
    logs: List[str] = []
    if not files:
        return ["[WARN] 업로드된 파일이 없습니다."]

    sh = open_sheet_by_env()

    for fname, raw in files.items():
        tab = _target_tab(fname)
        if not tab:
            logs.append(f"[SKIP] 파일명 규칙 불일치: {fname}")
            continue

        try:
            values = read_xlsx_values(raw, debug=True)
            if not values:
                logs.append(f"[ERROR] {tab}: {fname} 읽기 결과가 비어 있습니다.")
                continue
        except Exception as e:
            logs.append(f"[ERROR] {tab}: {fname} 읽기 실패 → {e}")
            continue

        try:
            _write_values_to_sheet(sh, tab, values, logs)
        except Exception as e:
            logs.append(f"[ERROR] {tab}: {fname} 반영 실패 → {e}")

    if not any(x.startswith("[OK]") for x in logs):
        logs.append("[WARN] 적용된 탭이 없습니다. 파일명 규칙/시트 권한을 확인하세요.")
    return logs


# ------------------------------------------------------
# Streamlit 업로더용 수집기 (기존 로직 유지)
# ------------------------------------------------------

def collect_xlsx_files(files) -> dict[str, BytesIO]:
    out: dict[str, BytesIO] = {}
    for f in files or []:
        name = getattr(f, "name", "")
        try:
            if name.lower().endswith('.xlsx'):
                out[name] = BytesIO(f.read())
        except Exception:
            # Streamlit 파일 객체가 이미 사용되었거나 읽기 실패 시 skip
            continue
    return out

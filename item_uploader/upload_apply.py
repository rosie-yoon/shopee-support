# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
from io import BytesIO
from typing import List, Optional, Tuple
from zipfile import ZipFile, ZIP_DEFLATED

import gspread
from gspread.utils import rowcol_to_a1
import pandas as pd
from openpyxl import load_workbook

# 프로젝트 공통 유틸
from .utils_common import open_sheet_by_env, safe_worksheet, with_retry

# --- [최종 수정] 엑셀 파일 파싱 안정화 로직 ---

def _sanitize_xlsx_for_openpyxl(file_bytes: bytes) -> BytesIO:
    """
    openpyxl의 특정 버전과 Shopee 엑셀 파일의 호환성 문제를 해결하기 위해,
    문제가 되는 sheetViews XML 블록을 제거하는 함수.
    """
    sanitized_buffer = BytesIO()
    try:
        with ZipFile(BytesIO(file_bytes), 'r') as zin, ZipFile(sanitized_buffer, 'w', ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                buffer = zin.read(item.filename)
                # 문제가 되는 worksheet XML 파일만 수정
                if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                    xml_content = buffer.decode('utf-8')
                    # sheetViews 블록을 통째로 제거 (가장 안정적인 방법)
                    cleaned_xml = re.sub(r'<sheetViews>.*?</sheetViews>', '', xml_content, flags=re.DOTALL)
                    buffer = cleaned_xml.encode('utf-8')
                zout.writestr(item, buffer)
        sanitized_buffer.seek(0)
        return sanitized_buffer
    except Exception:
        # 실패 시 원본 바이트를 그대로 반환
        return BytesIO(file_bytes)

def read_xlsx_values(bio: BytesIO) -> List[List[str]]:
    """
    업로드된 XLSX BytesIO를 안정적으로 파싱하여 2D 리스트로 반환합니다.
    1. 문제가 되는 XML 구조를 제거 (Sanitize).
    2. openpyxl로 보이는 행만 파싱 시도.
    3. 실패 또는 결과가 비정상적일 경우 pandas로 폴백.
    4. Shopee 메타 데이터 행 제거.
    """
    original_bytes = bio.getvalue()
    sanitized_bio = _sanitize_xlsx_for_openpyxl(original_bytes)

    # 1. openpyxl로 파싱 시도
    try:
        wb = load_workbook(sanitized_bio, data_only=True, read_only=True)
        ws = max(wb.worksheets, key=lambda s: (s.max_row or 0) * (s.max_column or 0))
        
        data = []
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if not (ws.row_dimensions.get(r_idx) and ws.row_dimensions[r_idx].hidden):
                str_row = [str(cell) if cell is not None else "" for cell in row]
                if any(str_row): # 빈 행이 아니면 추가
                    data.append(str_row)
    except Exception:
        data = []

    # 2. 결과가 비정상적일 경우 pandas로 폴백
    is_data_invalid = not data or (len(data) == 1 and len(data[0]) <= 1)
    if is_data_invalid:
        try:
            df = pd.read_excel(BytesIO(original_bytes), header=None, dtype=str).fillna('')
            data = df.values.tolist()
        except Exception:
            # pandas도 실패하면 빈 리스트 반환
            return []

    # 3. Shopee 메타 데이터 행 제거
    if not data:
        return []

    # 'et_title_'로 시작하는 헤더 행 제거
    while data and data[0] and any(str(c).startswith('et_title_') for c in data[0]):
        data.pop(0)
        
    # 'basic_info', 'media_info', 'sales_info' 행 제거
    while data and data[0] and str(data[0][0]) in ('basic_info', 'media_info', 'sales_info'):
        data.pop(0)

    return data


# ------------------------------------------------------
# Google Sheet 쓰기 (기존 코드와 동일)
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
# 파일명 → 탭 자동 라우팅 (기존 코드와 동일)
# ------------------------------------------------------
def _target_tab(filename: str) -> Optional[str]:
    low = filename.lower()
    if "basic" in low: return "BASIC"
    if "media" in low: return "MEDIA"
    if "sales" in low: return "SALES"
    return None


# ------------------------------------------------------
# 업로드 반영 엔트리 (기존 코드와 동일)
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
            values = read_xlsx_values(raw)
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
# Streamlit 업로더용 수집기 (기존 코드와 동일)
# ------------------------------------------------------
def collect_xlsx_files(files) -> dict[str, BytesIO]:
    out: dict[str, BytesIO] = {}
    for f in files or []:
        name = getattr(f, "name", "")
        if name.lower().endswith(".xlsx"):
            out[name] = BytesIO(f.read())
    return out


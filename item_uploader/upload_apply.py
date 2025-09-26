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
# XLSX 파싱 안정화 (Shopee/엑셀 시트뷰 불일치 + 숨김행/메타 안전 처리)
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
                    xml_content = buffer.decode('utf-8', errors='ignore')
                    # 1) <sheetViews ...> ... </sheetViews> 제거
                    cleaned_xml = re.sub(
                        r'<(?:\w+:)?sheetViews\b[^>]*>.*?</(?:\w+:)?sheetViews>',
                        '',
                        xml_content,
                        flags=re.DOTALL | re.IGNORECASE,
                    )
                    # 2) <pane .../> 단독 태그 제거
                    cleaned_xml = re.sub(
                        r'<(?:\w+:)?pane\b[^>]*/\s*>',
                        '',
                        cleaned_xml,
                        flags=re.DOTALL | re.IGNORECASE,
                    )
                    buffer = cleaned_xml.encode('utf-8')
                zout.writestr(item, buffer)
        sanitized_buffer.seek(0)
        return sanitized_buffer
    except Exception:
        return BytesIO(file_bytes)


# ------------------------
# robust openpyxl / pandas
# ------------------------
def _count_nonempty_cells(ws, max_rows: int = 800, max_cols: int = 256) -> int:
    """시트 앞쪽 일부에서 non-empty 셀 수를 세어 데이터가 가장 많은 시트를 고른다."""
    rows = min(ws.max_row or 0, max_rows)
    cols = min(ws.max_column or 0, max_cols)
    cnt = 0
    for r in ws.iter_rows(min_row=1, max_row=rows, min_col=1, max_col=cols, values_only=True):
        for c in r:
            if c not in (None, '', ' '):
                cnt += 1
    return cnt


def _read_with_openpyxl(sanitized_bio: BytesIO, logs: List[str], debug: bool = False) -> List[List[str]]:
    """정리된 바이트로 openpyxl을 이용해 값을 읽는다.
    - 숨김 행도 포함해서 읽고, 나중에 완전 빈 행만 제거
    - 모든 시트를 점수화하고 가장 데이터가 많은 시트를 선택
    """
    data: List[List[str]] = []
    try:
        sanitized_bio.seek(0)
        wb = load_workbook(sanitized_bio, data_only=True, read_only=True)
        scored = [(_count_nonempty_cells(s), s) for s in wb.worksheets]
        scored.sort(reverse=True, key=lambda x: x[0])

        if debug:
            logs.append("[DEBUG] openpyxl scores: " + ", ".join([f"{s.title}:{sc}" for sc, s in scored[:5]]))

        ws = (scored[0][1] if scored else wb.worksheets[0])
        if debug:
            logs.append(f"[DEBUG] openpyxl target sheet = {ws.title}")

        for row in ws.iter_rows(values_only=True):
            str_row = [str(cell) if cell is not None else "" for cell in row]
            data.append(str_row)

        data = [r for r in data if any(v.strip() for v in r)]
    except Exception as e:
        if debug:
            logs.append(f"[DEBUG] openpyxl read failed → {e}")
        data = []
    return data


def _read_with_pandas_fallback(sanitized_bio: BytesIO, logs: List[str], debug: bool = False) -> List[List[str]]:
    """pandas 폴백을 정리된 바이트로 수행. calamine 없이도 모든 시트를 스캔해 가장 데이터 많은 시트를 선택."""
    try:
        sanitized_bio.seek(0)
        df_all = pd.read_excel(
            sanitized_bio,
            header=None,
            dtype=str,
            sheet_name=None,      # 모든 시트
            engine="openpyxl"     # calamine 없이도 동작
        )
        best_df = None; best_score = -1; best_name = None
        for name, df in df_all.items():
            df = df.fillna('')
            score = int((df.values != '').sum())
            if score > best_score:
                best_score, best_df, best_name = score, df, name
        if best_df is not None:
            if debug:
                logs.append(f"[DEBUG] pandas used, target sheet={best_name}, score={best_score}")
            return best_df.values.tolist()
    except Exception as e:
        if debug:
            logs.append(f"[DEBUG] pandas fallback failed → {e}")
    return []


def read_xlsx_values(bio: BytesIO, logs: Optional[List[str]] = None, debug: bool = True) -> List[List[str]]:
    """
    업로드된 XLSX BytesIO를 안정적으로 파싱하여 2D 리스트로 반환합니다.
    1) XML sanitize
    2) openpyxl(숨김행 포함/모든 시트 점수화) → 비정상 시 pandas(all-sheets) 폴백
    3) Shopee 메타 행 최소 제거(최대 1회씩)
    """
    if logs is None:
        logs = []

    original_bytes = bio.getvalue()
    if len(original_bytes) < 1024:
        logs.append(f"[DEBUG] file too small: {len(original_bytes)} bytes")

    sanitized_bio = _sanitize_xlsx_for_openpyxl(original_bytes)

    # 1) openpyxl
    data = _read_with_openpyxl(sanitized_bio, logs, debug=debug)

    # 2) 비정상이면 pandas(all-sheets)
    if not data or (len(data) == 1 and len(data[0]) <= 1):
        data = _read_with_pandas_fallback(sanitized_bio, logs, debug=debug)
        if not data:
            if debug:
                logs.append("[DEBUG] both readers failed → return []")
            return []

    if debug:
        logs.append(f"[DEBUG] rows before meta-trim={len(data)}; head={data[:2]}")

    # 3) Shopee 메타 최소 제거 (최대 1회씩만)
    if data and data[0] and any(str(c).startswith('et_title_') for c in data[0]):
        data.pop(0); logs.append("[DEBUG] trimmed header row(et_title_*) once")
    if data and data[0] and str(data[0][0]).strip() in ('basic_info', 'media_info', 'sales_info'):
        data.pop(0); logs.append("[DEBUG] trimmed section label row once")

    # 최종 정리
    data = [r for r in data if any(v.strip() for v in r)]
    logs.append(f"[DEBUG] final rows={len(data)}; head={data[:3]}")
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
# 업로드 반영 엔트리 (logs를 read_xlsx_values에 넘기도록 수정)
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
            # ★ logs를 넘겨서 read_xlsx_values의 디버그를 UI 로그에 함께 표시
            values = read_xlsx_values(raw, logs, debug=True)
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
            continue
    return out

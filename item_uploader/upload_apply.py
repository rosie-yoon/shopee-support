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
# XLSX 파싱 안정화 (Shopee/엑셀 시트뷰 불일치 대응 클린본)
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


def _read_with_openpyxl(sanitized_bio: BytesIO) -> List[List[str]]:
    """정리된 바이트로 openpyxl을 이용해 보이는 값(data_only=True)만 읽는다."""
    data: List[List[str]] = []
    # openpyxl은 파일 포인터를 소비할 수 있으므로, 매 호출마다 0으로 이동
    try:
        sanitized_bio.seek(0)
        wb = load_workbook(sanitized_bio, data_only=True, read_only=True)
        # 가장 큰 시트를 선택 (행*열 기준)
        ws = max(wb.worksheets, key=lambda s: (s.max_row or 0) * (s.max_column or 0))

        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            # 숨김 행은 제외 (row_dimensions 없는 경우도 안전 처리)
            hidden = False
            try:
                dim = ws.row_dimensions.get(r_idx)
                hidden = bool(dim and getattr(dim, 'hidden', False))
            except Exception:
                hidden = False
            if hidden:
                continue

            str_row = [str(cell) if cell is not None else "" for cell in row]
            # 완전 빈 행은 제외
            if any(v != "" for v in str_row):
                data.append(str_row)
    except Exception:
        data = []
    return data


def _read_with_pandas_fallback(sanitized_bio: BytesIO) -> List[List[str]]:
    """pandas 폴백을 정리된 바이트로 수행. calamine가 있으면 우선 사용."""
    # 1) calamine 엔진 시도 (설치되어 있으면 openpyxl 의존 회피)
    try:
        sanitized_bio.seek(0)
        df = pd.read_excel(sanitized_bio, header=None, dtype=str, engine="calamine").fillna('')  # type: ignore[arg-type]
        return df.values.tolist()
    except Exception:
        pass

    # 2) 기본 엔진 (보통 openpyxl)로 재시도 — 이미 sanitize됐으므로 안전
    try:
        sanitized_bio.seek(0)
        df = pd.read_excel(sanitized_bio, header=None, dtype=str).fillna('')
        return df.values.tolist()
    except Exception:
        return []


def read_xlsx_values(bio: BytesIO) -> List[List[str]]:
    """
    업로드된 XLSX BytesIO를 안정적으로 파싱하여 2D 리스트로 반환합니다.
    1) XML sanitize로 sheetViews/pane 제거
    2) openpyxl로 시도 → 비정상 시 pandas로 폴백 (정리된 바이트 사용)
    3) Shopee 메타 데이터 행 제거
    """
    original_bytes = bio.getvalue()
    sanitized_bio = _sanitize_xlsx_for_openpyxl(original_bytes)

    # 1) openpyxl 경로
    data = _read_with_openpyxl(sanitized_bio)

    # 2) 결과가 비정상적이면 pandas 폴백 (정리된 바이트 기반)
    is_data_invalid = (not data) or (len(data) == 1 and len(data[0]) <= 1)
    if is_data_invalid:
        data = _read_with_pandas_fallback(sanitized_bio)
        if not data:
            return []

    # 3) Shopee 메타 데이터 행 제거
    #    - et_title_ 로 시작하는 헤더 제거
    while data and data[0] and any(str(c).startswith('et_title_') for c in data[0]):
        data.pop(0)
    #    - 구역 라벨(basic_info/media_info/sales_info) 제거
    while data and data[0] and str(data[0][0]) in ('basic_info', 'media_info', 'sales_info'):
        data.pop(0)

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
            values = read_xlsx_values(raw)
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

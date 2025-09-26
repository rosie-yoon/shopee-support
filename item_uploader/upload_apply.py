# -*- coding: utf-8 -*-
from __future__ import annotations
// ... existing code ...
def _read_with_openpyxl_visible_only(file_bytes: bytes) -> List[List[str]]:
    """
    openpyxl로 '보이는 행'만 읽는다. (열 숨김은 값 유지)
    오른쪽/아래쪽 연속 공백을 정리하여 2D 배열 균일화.
    """
    from openpyxl import load_workbook

    # 💥💥💥 [BUG FIX] 
    # Shopee에서 생성된 엑셀 파일의 특정 메타데이터와 openpyxl 간의 호환성 문제를 피하기 위해
    # `load_workbook` 호출 시 `keep_vba=False`와 같은 옵션 대신,
    # 아래 `_read_with_pandas_all_rows` 함수를 우선적으로 사용하도록 `read_xlsx_values` 로직을 변경합니다.
    # 이 함수 자체는 유지하되, 호출 흐름을 변경하여 안정성을 확보합니다.

    wb = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True, keep_links=False)
    # 데이터가 가장 많은 시트를 선택 (필요시 wb.worksheets[0]으로 고정 가능)
    target = max(wb.worksheets, key=lambda s: (s.max_row or 0) * (s.max_column or 0))

// ... existing code ...
def read_xlsx_values(bio: BytesIO) -> List[List[str]]:
    """
    업로드된 XLSX BytesIO → 2D values
      - sheetViews/pane 제거(안정화)
      - (수정) 안정성을 위해 pandas 파서를 우선적으로 시도
      - Shopee 라벨/메타 행 제거
    """
    sanitized = _sanitize_xlsx_remove_sheetviews(bio)
    sanitized.seek(0)
    raw_bytes = sanitized.read()

    # --- 💥💥💥 [BUG FIX] 수정된 로직 ---
    # 1. 안정적인 pandas로 먼저 파싱을 시도합니다.
    try:
        values = _read_with_pandas_all_rows(raw_bytes)
    except Exception:
        values = []

    # 2. pandas 결과가 비정상적일 경우에만 openpyxl을 폴백으로 사용합니다.
    is_invalid = (len(values) == 0) or (len(values) <= 1 and (len(values[0]) if values else 0) <= 1)
    if is_invalid:
        try:
            openpyxl_values = _read_with_openpyxl_visible_only(raw_bytes)
            # openpyxl 결과가 더 나은 경우에만 교체
            if len(openpyxl_values) > len(values) or (openpyxl_values and (not values or len(openpyxl_values[0]) > len(values[0]))):
                values = openpyxl_values
        except Exception:
            pass # openpyxl 실패는 무시

    # 3. 마지막으로 메타데이터 행을 제거합니다.
    final_values = _strip_shopee_meta_rows(values)
    return final_values


# ------------------------------------------------------
# 5) Google Sheet 쓰기 (RAW + 청크)
#    - .env: UPLOAD_CHUNK_ROWS (0이면 한 번에)
// ... existing code ...

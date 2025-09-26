# -*- coding: utf-8 -*-
"""
automation_steps.py (CLEAN, final)
- Step 1 ~ Step 7 자동화 로직
- ref(참조 시트)가 None이더라도 안전하게 열리도록 보강
- TemplateDict 탭 누락/권한 이슈 시 명확한 예외 메시지
"""

from __future__ import annotations
import re
import os
from io import BytesIO
from collections import defaultdict
from typing import Dict, List, Tuple, Optional, Set

import gspread
from gspread.cell import Cell
from gspread.utils import rowcol_to_a1
from gspread.exceptions import WorksheetNotFound
import pandas as pd

from .utils_common import (
    load_env, with_retry, safe_worksheet, header_key, top_of_category,
    get_tem_sheet_name, get_env, get_bool_env, hex_to_rgb01, strip_category_id,
    open_ref_by_env, 
)

# ==============================================================================
# 공통 헬퍼
# ==============================================================================

def _ensure_ref(ref_obj: Optional[gspread.Spreadsheet]) -> gspread.Spreadsheet:
    """ref가 None이면 환경설정(REFERENCE_SHEET_KEY)로 열어 보장."""
    return ref_obj or open_ref_by_env()

def _pick_index_by_candidates(header_row: List[str], candidates: List[str]) -> int:
    """헤더 행에서 후보명(정규화)으로 가장 그럴듯한 인덱스 찾기 (정확 > 부분 일치)"""
    keys = [header_key(x) for x in header_row]
    # 정확 일치
    for cand in candidates:
        ck = header_key(cand)
        for i, k in enumerate(keys):
            if k == ck:
                return i
    # 부분 일치
    for cand in candidates:
        ck = header_key(cand)
        if not ck:
            continue
        for i, k in enumerate(keys):
            if ck in k:
                return i
    return -1

def _find_col_index(keys: List[str], name: str, extra_alias: List[str]=[]) -> int:
    """헤더 키 목록(keys=header_key 적용된 리스트)에서 name 또는 alias를 찾음"""
    tgt = header_key(name)
    aliases = [header_key(a) for a in extra_alias] + [tgt]
    # 정확 매칭
    for i, k in enumerate(keys):
        if k in aliases:
            return i
    # 포함 매칭
    for i, k in enumerate(keys):
        if any(a and a in k for a in aliases):
            return i
    return -1

def _append_failures(sh, rows: List[List[str]]):
    """Failures 탭에 rows를 append. 공간 부족 시 자동 resize."""
    if not rows:
        return
    try:
        ws = safe_worksheet(sh, "Failures")
        vals = with_retry(lambda: ws.get_all_values()) or []
        start_row = len(vals) + 1
        end_row = start_row + len(rows) - 1

        if end_row > ws.row_count:
            with_retry(lambda: ws.resize(rows=end_row + 100, cols=max(ws.col_count, 10)))

        with_retry(lambda: ws.update(values=rows, range_name=f"A{start_row}"))
    except WorksheetNotFound:
        ws = with_retry(lambda: sh.add_worksheet(title="Failures", rows=1000, cols=10))
        header = [["PID","Category","Name","Reason","Detail"]]
        with_retry(lambda: ws.update(values=header + rows, range_name="A1"))

# ==============================================================================
# STEP 1: TEM_OUTPUT 생성
# ==============================================================================

def run_step_1(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet):
    """Step 1: BASIC+MEDIA -> TEM_OUTPUT 생성 (+ SALES로 SKU/Parent SKU 매핑)"""
    print("\n[ Automation ] Starting Step 1: Build TEM_OUTPUT...")

    # ✅ 전달된 ref 인자를 신뢰하지 않고 항상 '참조 시트'를 강제 오픈
    ref = open_ref_by_env()

    # 어떤 문서를 보고 있는지 확인 (디버그)
    sh_id  = getattr(sh, "id", None);  sh_url  = getattr(sh, "url", "(n/a)")
    ref_id = getattr(ref, "id", None); ref_url = getattr(ref, "url", "(n/a)")
    print(f"[STEP1][MAIN] id={sh_id} url={sh_url}")
    print(f"[STEP1][REF ] id={ref_id} url={ref_url}")
    if sh_id and ref_id and sh_id == ref_id:
        raise RuntimeError(
            "[STEP1] 참조 시트가 메인 시트와 동일합니다. "
            "⇒ REFERENCE_SHEET_KEY가 메인 시트 키로 설정되었는지 확인하세요."
        )

    basic_header = int(get_env("BASIC_HEADER_ROW", "2"))
    basic_first  = int(get_env("BASIC_FIRST_DATA_ROW", "3"))
    media_header = int(get_env("MEDIA_HEADER_ROW", "2"))
    media_first  = int(get_env("MEDIA_FIRST_DATA_ROW", "6"))
    tem_name     = get_tem_sheet_name()

    # BASIC / MEDIA 읽기
    basic_ws = safe_worksheet(sh, "BASIC")
    media_ws = safe_worksheet(sh, "MEDIA")
    basic_vals = with_retry(lambda: basic_ws.get_all_values()) or []
    media_vals = with_retry(lambda: media_ws.get_all_values()) or []
    if len(basic_vals) < basic_header or len(media_vals) < media_header:
        print("[!] BASIC or MEDIA 시트가 비어 있습니다.")
        return

    # ✅ TemplateDict: 참조 시트에서만 엄격 매칭 (폴백 없음)
    REF_TEMPLATE_TAB = "TemplateDict"
    ref_titles = [ws.title for ws in with_retry(lambda: ref.worksheets())]
    print(f"[STEP1][REF] tabs={ref_titles}")
    if REF_TEMPLATE_TAB not in ref_titles:
        raise RuntimeError(
            f"[STEP1] 참조 시트는 열렸지만 '{REF_TEMPLATE_TAB}' 탭이 없습니다. "
            f"실제 탭들={ref_titles}"
        )

    template_dict_ws = safe_worksheet(ref, REF_TEMPLATE_TAB)
    print(f"[STEP1] Using TemplateDict worksheet title = '{template_dict_ws.title}'")

    template_vals = with_retry(lambda: template_dict_ws.get_all_values()) or []
    if not template_vals or len(template_vals) < 2:
        raise RuntimeError("[STEP1] TemplateDict 탭이 비어 있거나 유효한 헤더/데이터가 없습니다.")

    template_dict = {
        header_key(row[0]): [str(x or "").strip() for x in row[1:]]
        for row in template_vals[1:] if (row[0] or "").strip()
    }

# ==============================================================================
# STEP 2: Mandatory 기본값 채우기 (+ 색칠)
# ==============================================================================

def run_step_2(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet):
    """Step 2: TEM_OUTPUT에 Mandatory 기본값 채우기 + 색칠"""
    print("\n[ Automation ] Starting Step 2: Fill Mandatory Defaults...")
    ref = _ensure_ref(ref)

    tem_name = get_tem_sheet_name()
    color_hex = get_env("COLOR_HEX_MANDATORY", "#FFF9C4")
    overwrite = get_bool_env("OVERWRITE_NONEMPTY", False)

    try:
        tem_ws = safe_worksheet(sh, tem_name)
    except WorksheetNotFound:
        print(f"[!] {tem_name} 탭 없음. Step1 선행 필요."); return

    def _read_defaults_ws(ws):
        vals = with_retry(lambda: ws.get_all_values()) or []
        if not vals: return {}
        keys = [header_key(x) for x in vals[0]]
        c_idx = _find_col_index(keys, "category")
        a_idx = _find_col_index(keys, "attribute", ["attr", "property"])
        d_idx = _find_col_index(keys, "defaultvalue", ["default"])
        if min(c_idx, a_idx, d_idx) < 0: return {}
        out = {}
        for r in range(1, len(vals)):
            row = vals[r]
            cat  = (row[c_idx] if c_idx < len(row) else "").strip()
            attr = (row[a_idx] if a_idx < len(row) else "").strip()
            dval = (row[d_idx] if d_idx < len(row) else "").strip()
            if cat and attr:
                out.setdefault((cat or "").strip().lower(), {})[header_key(attr)] = dval
        return out

    # 레퍼런스의 MandatoryDefaults_* 시트들 합치기
    sheets = with_retry(lambda: ref.worksheets())
    defaults_map: Dict[str, Dict[str, str]] = {}
    for ws in sheets:
        if ws.title.lower().startswith("mandatorydefaults_"):
            for k, d in _read_defaults_ws(ws).items():
                defaults_map.setdefault(k, {}).update(d)

    # Category에서 Mandatory로 표시된 헤더 수집
    cat_props_ws = safe_worksheet(ref, get_env("CAT_PROPS_SHEET", "cat props"))
    cat_props_vals = with_retry(lambda: cat_props_ws.get_all_values()) or []
    catprops_map: Dict[str, List[str]] = {}
    if cat_props_vals:
        hdr_keys = [header_key(x) for x in cat_props_vals[0]]
        for r in range(1, len(cat_props_vals)):
            row = cat_props_vals[r]
            cat_raw = (row[0] if len(row) > 0 else "").strip()
            if not cat_raw: continue
            mand_list = [hdr_keys[j] for j, cell in enumerate(row) if str(cell).strip().lower() == "mandatory"]
            if mand_list:
                catprops_map[(cat_raw or "").strip().lower()] = mand_list

    vals = with_retry(lambda: tem_ws.get_all_values()) or []
    if not vals: print("[!] TEM_OUTPUT 비어 있음."); return

    # 색칠/업데이트를 위한 sheetId
    meta = with_retry(lambda: sh.fetch_sheet_metadata())
    sheet_id = next((s["properties"]["sheetId"]
                     for s in meta["sheets"]
                     if s["properties"]["title"] == tem_name), None)
    if sheet_id is None:
        print("[!] 시트 ID 찾지 못함."); return

    updates: List[Cell] = []
    color_ranges_by_col = defaultdict(list)
    current_hdr_keys: Optional[List[str]] = None
    total_filled = 0

    for r0, row in enumerate(vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_hdr_keys = [header_key(h) for h in row[1:]]
            continue
        if not current_hdr_keys: 
            continue

        pid = (row[0] if len(row) > 0 else "").strip()
        cat_raw = (row[1] if len(row) > 1 else "").strip()
        if not pid or not cat_raw:
            continue
        norm_cat = (cat_raw or "").strip().lower()

        # 색칠(카테고리별 mandatory 헤더)
        if norm_cat in catprops_map:
            for attr_norm in catprops_map[norm_cat]:
                j = _find_col_index(current_hdr_keys, attr_norm)
                if j >= 0:
                    color_ranges_by_col[j].append((r0, r0 + 1))

        # 기본값 채우기
        if norm_cat in defaults_map:
            for attr_norm, dval in defaults_map[norm_cat].items():
                if not dval: 
                    continue
                j = _find_col_index(current_hdr_keys, attr_norm)
                if j < 0: 
                    continue
                col_1based = j + 2
                cur = (row[col_1based - 1] if len(row) >= col_1based else "").strip()
                if not cur or overwrite:
                    updates.append(Cell(row=r0 + 1, col=col_1based, value=dval))
                    total_filled += 1

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="RAW"))

    # 색칠 요청 머지 유틸
    def _merge(spans):
        if not spans: return []
        spans.sort(); merged = [spans[0]]
        for s, e in spans[1:]:
            ls, le = merged[-1]
            if s <= le: merged[-1] = (ls, max(le, e))
            else: merged.append((s, e))
        return merged

    requests = []
    color = hex_to_rgb01(color_hex)
    for j, spans in color_ranges_by_col.items():
        for s, e in _merge(spans):
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": s, "endRowIndex": e,
                        "startColumnIndex": 1 + j, "endColumnIndex": 1 + j + 1
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": color}},
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })

    if requests:
        with_retry(lambda: sh.batch_update({"requests": requests}))

    print("========== STEP 2 RESULT ==========")
    print(f"채워진 셀 수: {total_filled:,}")
    print(f"색칠된 열 개수: {len(color_ranges_by_col):,}")
    print("Step 2: Fill Mandatory Defaults Finished.")

# ==============================================================================
# STEP 3: FDA 코드 채우기
# ==============================================================================

def run_step_3(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet, overwrite: bool = False):
    """Reference의 대상 카테고리 목록을 기준으로 TEM_OUTPUT에 고정 FDA 코드 채우기"""
    print("\n[ Automation ] Starting Step 3: Fill FDA Code...")
    ref = _ensure_ref(ref)

    tem_name = get_tem_sheet_name()
    fda_sheet_name = get_env("FDA_CATEGORIES_SHEET_NAME", "TH Cos")
    fda_header = get_env("FDA_HEADER_NAME", "FDA Registration No.")
    FDA_CODE = "10-1-9999999"

    # Reference 시트에서 FDA 대상 카테고리 목록 읽기
    try:
        fda_ws = safe_worksheet(ref, fda_sheet_name)
        fda_vals_2d = with_retry(lambda: fda_ws.get_values('A:A', value_render_option='UNFORMATTED_VALUE'))
        fda_vals = [r[0] for r in (fda_vals_2d or []) if r and str(r[0]).strip()]
        target_categories = {str(cat).strip().lower() for cat in fda_vals if str(cat).strip()}
    except Exception as e:
        print(f"[!] '{fda_sheet_name}' 탭을 읽는 데 실패했습니다: {e}. Step 3을 건너뜁니다.")
        return

    # TEM_OUTPUT 읽기
    try:
        tem_ws = safe_worksheet(sh, tem_name)
        vals = with_retry(lambda: tem_ws.get_all_values()) or []
    except WorksheetNotFound:
        print(f"[!] {tem_name} 탭 없음. Step1 선행 필요."); return
    if not vals:
        print("[!] TEM_OUTPUT 비어 있음."); return

    updates: List[Cell] = []
    current_keys: Optional[List[str]] = None
    col_category_B = -1
    col_fda_B = -1
    updated_rows = 0

    for r0, row in enumerate(vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_keys = [header_key(h) for h in row[1:]]
            col_category_B = _find_col_index(current_keys, "category")
            col_fda_B = _find_col_index(current_keys, fda_header)
            continue
        if not current_keys or col_fda_B < 0 or col_category_B < 0:
            continue

        pid = (row[0] if len(row) > 0 else "").strip()
        if not pid: 
            continue

        category_val_raw = (row[col_category_B + 1] if len(row) > (col_category_B + 1) else "").strip()
        category_val_normalized = category_val_raw.lower()

        if category_val_normalized and category_val_normalized in target_categories:
            c_fda_sheet_col = col_fda_B + 2
            cur_fda = (row[c_fda_sheet_col - 1] if len(row) >= c_fda_sheet_col else "").strip()
            if not cur_fda or overwrite:
                updates.append(Cell(row=r0 + 1, col=c_fda_sheet_col, value=FDA_CODE))
                updated_rows += 1

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="RAW"))

    print("========== STEP 3 FDA RESULT (WRITE) ==========")
    print(f"적용된 셀 수: {updated_rows:,}")
    print("Step 3: Fill FDA Code Finished.")

# ==============================================================================
# STEP 4: 기타 필드 채우기 (Stock / Days to ship / Weight / Brand)
# ==============================================================================

def run_step_4(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet):
    print("\n[ Automation ] Starting Step 4: Fill Other Fields...")
    ref = _ensure_ref(ref)

    tem_name = get_tem_sheet_name()
    STOCK_VALUE = int(get_env("STEP4_STOCK_VALUE", "1000"))
    DTOS_VALUE  = int(get_env("STEP4_DTOS_VALUE", "1"))

    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []
    if not tem_vals:
        print("[!] TEM_OUTPUT 비어 있음."); return

    # 보조 데이터
    try:
        margin_ws = safe_worksheet(sh, "MARGIN")
        margin_vals = with_retry(lambda: margin_ws.get_all_values()) or []
    except Exception:
        margin_vals = []

    try:
        brand_ws = safe_worksheet(ref, "Brand")
        brand_vals = with_retry(lambda: brand_ws.get_all_values()) or []
    except Exception:
        brand_vals = []

    sku_to_weight: Dict[str, str] = {}
    sku_to_brand_name: Dict[str, str] = {}
    if margin_vals:
        mh = margin_vals[0]
        idx_sku   = _pick_index_by_candidates(mh, ["sku", "seller_sku"])
        idx_brandn= _pick_index_by_candidates(mh, ["brand", "brand name"])
        idx_wgt   = _pick_index_by_candidates(mh, ["weight", "package weight"])
        if idx_sku >= 0:
            for r in range(1, len(margin_vals)):
                row = margin_vals[r]
                sku = (row[idx_sku] if idx_sku < len(row) else "").strip()
                if not sku: continue
                if 0 <= idx_wgt    < len(row): sku_to_weight[sku]     = (row[idx_wgt] or "").strip()
                if 0 <= idx_brandn < len(row): sku_to_brand_name[sku] = (row[idx_brandn] or "").strip()

    brand_name_to_code: Dict[str, str] = {}
    if brand_vals and len(brand_vals[0]) >= 3:
        for r in range(1, len(brand_vals)):
            row = brand_vals[r]
            if len(row) < 3: continue
            bname = (row[1] or "").strip()
            bcode = (row[2] or "").strip()
            if bname: brand_name_to_code[re.sub(r"\s+", " ", bname.lower())] = bcode

    failures: List[List[str]] = []
    cells_to_update: List[Cell] = []
    cnt_stock = cnt_dtos = cnt_weight = cnt_brand = 0
    current_headers: Optional[List[str]] = None
    idx_stock_B = idx_dtos_B = idx_weight_B = idx_brand_B = idx_sku_B = -1

    for r in range(len(tem_vals)):
        row = tem_vals[r]
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = row[1:]
            hdr_keys = [header_key(h) for h in current_headers]
            idx_stock_B  = _find_col_index(hdr_keys, "stock")
            idx_dtos_B   = _find_col_index(hdr_keys, "daystoship")
            idx_weight_B = _find_col_index(hdr_keys, "weight")
            idx_brand_B  = _find_col_index(hdr_keys, "brand")
            idx_sku_B    = _find_col_index(hdr_keys, "sku")
            continue
        if not current_headers:
            continue

        pid = (row[0] if len(row) > 0 else "").strip()

        if idx_stock_B >= 0:
            c = idx_stock_B + 2
            if (row[c - 1] if len(row) >= c else "") != str(STOCK_VALUE):
                cells_to_update.append(Cell(row=r + 1, col=c, value=str(STOCK_VALUE)))
                cnt_stock += 1

        if idx_dtos_B >= 0:
            c = idx_dtos_B + 2
            if (row[c - 1] if len(row) >= c else "") != str(DTOS_VALUE):
                cells_to_update.append(Cell(row=r + 1, col=c, value=str(DTOS_VALUE)))
                cnt_dtos += 1

        sku_val = ""
        if idx_sku_B >= 0:
            csku = idx_sku_B + 2
            sku_val = (row[csku - 1] if len(row) >= csku else "").strip()

        if idx_weight_B >= 0 and sku_val:
            w = sku_to_weight.get(sku_val)
            if w:
                c = idx_weight_B + 2
                if (row[c - 1] if len(row) >= c else "") != w:
                    cells_to_update.append(Cell(row=r + 1, col=c, value=w))
                    cnt_weight += 1
            else:
                failures.append([pid, "", "", "WEIGHT_MAP_MISSING", f"sku={sku_val}"])

        if idx_brand_B >= 0 and sku_val:
            bname = sku_to_brand_name.get(sku_val)
            bcode = brand_name_to_code.get(re.sub(r"\s+", " ", bname.lower())) if bname else None
            new_bcode = bcode if bcode else "0"
            c = idx_brand_B + 2
            if (row[c - 1] if len(row) >= c else "") != new_bcode:
                cells_to_update.append(Cell(row=r + 1, col=c, value=new_bcode))
                cnt_brand += 1
            if bname and not bcode:
                failures.append([pid, "", "", "BRAND_CODE_NOT_FOUND", f"brand_name={bname}"])

    if cells_to_update:
        with_retry(lambda: tem_ws.update_cells(cells_to_update, value_input_option="USER_ENTERED"))
    if failures:
        _append_failures(sh, failures)

    print("========== STEP 4 RESULT ==========")
    print(f"Stock/DTOS/Weight/Brand 채움: {cnt_stock}/{cnt_dtos}/{cnt_weight}/{cnt_brand}")
    print("Step 4: Fill Other Fields Finished.")

# ==============================================================================
# STEP 5: 기타 필수정보 채우기 (Desc / VariationIntegration / GlobalSKUPrice)
# ==============================================================================

def run_step_5(sh: gspread.Spreadsheet):
    print("\n[ Automation ] Starting Step 5: Fill essential info...")

    tem_name = get_tem_sheet_name()
    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []

    basic_ws = safe_worksheet(sh, "BASIC")
    basic_vals = with_retry(lambda: basic_ws.get_all_values()) or []

    margin_ws = safe_worksheet(sh, "MARGIN")
    margin_vals = with_retry(lambda: margin_ws.get_all_values()) or []

    # 데이터 맵
    pid_to_desc = {row[0].strip(): (row[3] if len(row) > 3 else "") for row in basic_vals[1:] if row and row[0].strip()}
    sku_to_price = {row[0].strip(): (row[4] if len(row) > 4 else "") for row in margin_vals[1:] if row and row[0].strip()}

    updates: List[Cell] = []
    current_headers: Optional[List[str]] = None
    pid_groups = defaultdict(list)
    idx_desc = idx_var_integ = idx_price = idx_sku = -1  # ← 초기화(희귀 케이스 방지)

    for r_idx, row in enumerate(tem_vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = [header_key(h) for h in row[1:]]
            idx_desc      = _find_col_index(current_headers, "productdescription")
            idx_var_integ = _find_col_index(current_headers, "variationintegration")
            idx_price     = _find_col_index(current_headers, "globalskuprice")
            idx_sku       = _find_col_index(current_headers, "sku")
            continue
        if not current_headers:
            continue

        pid = (row[0] if len(row) > 0 else "").strip()
        if not pid:
            continue

        pid_groups[pid].append(r_idx + 1)

        # 1. Description
        if idx_desc != -1:
            desc = pid_to_desc.get(pid, "")
            updates.append(Cell(row=r_idx + 1, col=idx_desc + 2, value=desc))

        # 3. Global SKU Price
        if idx_price != -1 and idx_sku != -1:
            sku_val = (row[idx_sku + 1] if len(row) > idx_sku + 1 else "").strip()
            if sku_val:
                price = sku_to_price.get(sku_val, "")
                updates.append(Cell(row=r_idx + 1, col=idx_price + 2, value=price))

    # 2. Variation Integration
    if idx_var_integ != -1:
        for pid, rows in pid_groups.items():
            if len(rows) > 1:  # Only for variations
                v_code = f"V{pid}"
                for r in rows:
                    updates.append(Cell(row=r, col=idx_var_integ + 2, value=v_code))

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="USER_ENTERED"))

    print("Step 5: Fill essential info Finished.")

# ==============================================================================
# STEP 6: Cover Image URL 생성 (Parent SKU 우선 규칙)
# ==============================================================================

def run_step_6(sh: gspread.Spreadsheet, shop_code: str):
    print("\n[ Automation ] Starting Step 6: Generate Cover Image URLs...")

    tem_name = get_tem_sheet_name()
    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []

    host = get_env("IMAGE_HOSTING_URL", "")
    if not host.endswith("/"):
        host += "/"

    updates: List[Cell] = []
    current_headers: Optional[List[str]] = None
    idx_cover = idx_sku = idx_psku = -1

    for r_idx, row in enumerate(tem_vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = [header_key(h) for h in row[1:]]
            idx_cover = _find_col_index(current_headers, "coverimage")
            idx_sku   = _find_col_index(current_headers, "sku")
            idx_psku  = _find_col_index(current_headers, "parentsku")
            continue
        if not current_headers or idx_cover == -1:
            continue

        psku_val = (row[idx_psku + 1] if idx_psku != -1 and len(row) > idx_psku + 1 else "").strip()
        sku_val  = (row[idx_sku  + 1] if idx_sku  != -1 and len(row) > idx_sku  + 1 else "").strip()

        sku_for_url = psku_val if psku_val else sku_val
        if sku_for_url:
            url = f"{host}{sku_for_url}_C_{shop_code}.jpg"
            updates.append(Cell(row=r_idx + 1, col=idx_cover + 2, value=url))

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="USER_ENTERED"))

    print("Step 6: Generate Cover Image URLs Finished.")

# ==============================================================================
# STEP 7: 최종 템플릿 분할 & 다운로드
# ==============================================================================

def run_step_7(sh: gspread.Spreadsheet):
    """
    TEM_OUTPUT을 TopLevel Category 단위로 분할하여 엑셀 파일 생성
    - Sheets 429(쿼터) 회피: with_retry에 넉넉한 백오프 적용
    - get_values() 남용 방지: 범위를 지정해 읽기
    """
    import re
    from io import BytesIO
    import pandas as pd
    from gspread.utils import rowcol_to_a1

    print("\n[ Automation ] Starting Step 7: Generating final template file...")

    tem_name = get_tem_sheet_name()
    tem_ws = safe_worksheet(sh, tem_name)

    # 1) 헤더 1행 먼저 읽고, 대략적인 범위를 계산해 한 번에 읽기
    header = with_retry(lambda: tem_ws.row_values(1), retries=8, base_delay=3.0, backoff=2.0, max_delay=70.0)
    if not header:
        print("[!] TEM_OUTPUT header is empty. Cannot generate file.")
        return None

    max_cols = max(1, len(header))
    # 행 수는 넉넉히 20000까지 (필요시 ENV로 튜닝 가능)
    max_rows = 20000
    rng = f"A1:{rowcol_to_a1(max_rows, max_cols)}"

    all_data = with_retry(lambda: tem_ws.get(rng), retries=8, base_delay=3.0, backoff=2.0, max_delay=70.0)
    if not all_data:
        print("[!] TEM_OUTPUT sheet is empty within range. Cannot generate file.")
        return None

    # 2) DataFrame 구성
    df = pd.DataFrame(all_data)
    # 최소 2열은 있다고 가정 (1열: PID, 2열: Category/Headers 시작)
    if df.shape[1] < 2:
        print("[!] TEM_OUTPUT has insufficient columns.")
        return None

    # 'category' 헤더가 있는 행 인덱스 탐지 (두번째 컬럼 기준)
    # 일부 행은 빈 값일 수 있으니, 문자열로 변환 후 비교
    col1 = df.iloc[:, 1].astype(str).str.lower()
    header_indices = col1.index[col1.eq('category')]
    if len(header_indices) == 0:
        print("[!] No valid header rows found in TEM_OUTPUT.")
        return None

    # 3) 분할 저장
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for i, header_idx in enumerate(header_indices):
            start_row = header_idx + 1
            end_row = header_indices[i + 1] if i + 1 < len(header_indices) else len(df)

            chunk_df = df.iloc[start_row:end_row]
            if chunk_df.empty:
                continue

            # 첫 데이터 행의 카테고리로 Top Level 판단
            try:
                first_cat = str(chunk_df.iloc[0, 1] or "")
            except Exception:
                first_cat = "UNKNOWN"

            top_level_name = top_of_category(first_cat) or "UNKNOWN"
            # 시트명 제약 문자 제거 + 31자 제한
            sheet_name = re.sub(r'[\\s/\\\\*?:\\[\\]]', '_', top_level_name.title())[:31] or "Sheet1"

            # 첫 컬럼(Product ID) 제거 → B열부터 저장
            sub_df = chunk_df.iloc[:, 1:].copy()

            # SKU(첫 열) 공백-하이픈 정리: 'ABC - 123' -> 'ABC-123'
            try:
                sub_df.iloc[:, 0] = sub_df.iloc[:, 0].astype(str).str.replace(r'\s*-\s*', '-', regex=True)
            except Exception:
                pass

            # 헤더 없는 raw로 저장 (원본 템플릿 포맷 유지)
            sub_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    output.seek(0)
    print("Step 7: Final template file generated successfully.")
    return output

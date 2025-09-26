# -*- coding: utf-8 -*-
"""
automation_steps_revised.py
- Step 1부터 Step 7까지의 핵심 자동화 로직을 통합한 모듈입니다.
- 각 함수는 main_controller.py에서 호출됩니다.
"""

from __future__ import annotations
import re
import os
import random
import hashlib
from collections import defaultdict
from typing import Dict, List, Tuple, Optional, Set
from io import BytesIO

import gspread
from gspread.cell import Cell
from gspread.utils import rowcol_to_a1
from gspread.exceptions import WorksheetNotFound
import pandas as pd


from .utils_common import (
    load_env, with_retry, safe_worksheet, header_key, top_of_category,
    get_tem_sheet_name, get_env, get_bool_env, hex_to_rgb01, strip_category_id
)

# ==============================================================================
# 공통 헬퍼 함수
# ==============================================================================

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
         # 시트가 아예 없는 경우 새로 만듦 (main_controller에서 초기화하지만 안전장치)
        ws = with_retry(lambda: sh.add_worksheet(title="Failures", rows=1000, cols=10))
        header = [["PID","Category","Name","Reason","Detail"]]
        with_retry(lambda: ws.update(values=header + rows, range_name="A1"))


# ==============================================================================
# STEP 1: TEM_OUTPUT 생성
# ==============================================================================

def run_step_1(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet):
    """Step 1: BASIC+MEDIA -> TEM_OUTPUT 생성 (+ SALES로 SKU/Parent SKU 매핑)"""
    print("\n[ Automation ] Starting Step 1: Build TEM_OUTPUT...")

    basic_header = int(get_env("BASIC_HEADER_ROW", "2"))
    basic_first  = int(get_env("BASIC_FIRST_DATA_ROW", "3"))
    media_header = int(get_env("MEDIA_HEADER_ROW", "2"))
    media_first  = int(get_env("MEDIA_FIRST_DATA_ROW", "6"))
    ref_sheet = get_env("TEMPLATE_DICT_SHEET_NAME", "TemplateDict")
    tem_name = get_tem_sheet_name()

    basic_ws = safe_worksheet(sh, "BASIC")
    media_ws = safe_worksheet(sh, "MEDIA")
    basic_vals = with_retry(lambda: basic_ws.get_all_values())
    media_vals = with_retry(lambda: media_ws.get_all_values())

    if len(basic_vals) < basic_header or len(media_vals) < media_header:
        print("[!] BASIC or MEDIA 시트가 비어 있습니다.")
        return

    template_dict_ws = safe_worksheet(ref, ref_sheet)
    template_vals = with_retry(lambda: template_dict_ws.get_all_values()) or []
    template_dict = {
        header_key(row[0]): [str(x or "").strip() for x in row[1:]]
        for row in template_vals[1:] if (row[0] or "").strip()
    }

    class MediaHeader:
        def __init__(self):
            self.pid = -1; self.pname = -1; self.category = -1; self.cover = -1
            self.item_images: List[int] = []; self.var_label = -1
            self.opt_name_cols: List[int] = []; self.opt_img_cols: List[int] = []

    def parse_media_header_row(header_row: List[str]) -> MediaHeader:
        h = MediaHeader()
        keys = [header_key(x) for x in header_row]
        h.pid      = _find_col_index(keys, "productid",["pid","itemid","ettitleproductid"])
        h.pname    = _find_col_index(keys, "productname", ["itemname","name"])
        h.category = _find_col_index(keys, "category")
        h.cover    = _find_col_index(keys, "coverimage", ["coverimg"])
        h.var_label= _find_col_index(keys, "variationname1", ["variationname","variation"])
        for i, raw in enumerate(header_row):
            if header_key(raw).startswith("itemimage"): h.item_images.append(i)
        patt = re.compile(r"^option(\d+)name$")
        for idx, raw in enumerate(header_row):
            m = patt.match(header_key(raw))
            if not m: continue
            n = m.group(1)
            h.opt_name_cols.append(idx)
            img_idx = next((j for j, r in enumerate(header_row) if header_key(r)==f"option{n}image"), -1)
            h.opt_img_cols.append(img_idx)
        return h

    media_hdr = parse_media_header_row(media_vals[media_header - 1])

    parent_sku_map: dict[str, str] = {}
    sku_by_pid_opt: dict[tuple[str, str], str] = {}
    try:
        sales_ws = safe_worksheet(sh, "SALES")
        sales_vals = with_retry(lambda: sales_ws.get_all_values()) or []
        if sales_vals:
            hdr = sales_vals[0]
            pid_idx = _pick_index_by_candidates(hdr, ["product id","pid","item id","et_title_product_id"])
            psku_idx = _pick_index_by_candidates(hdr, ["parent sku","parent_sku","seller sku","seller_sku","et_title_parent_sku"])
            var_name_idx = _pick_index_by_candidates(hdr, ["variation name","option name","option 1 name","variation option","variation","option"])
            sku_idx = _pick_index_by_candidates(hdr, ["sku","variation sku","child sku","option sku","seller_child_sku","et_title_child_sku"])
            
            if pid_idx >= 0:
                for r in range(1, len(sales_vals)):
                    row = sales_vals[r]
                    pid = (row[pid_idx] if pid_idx < len(row) else "").strip()
                    if not pid: continue
                    if 0 <= psku_idx < len(row) and (row[psku_idx] or "").strip():
                        parent_sku_map[pid] = row[psku_idx].strip()
                    if 0 <= var_name_idx < len(row) and 0 <= sku_idx < len(row):
                        vname = (row[var_name_idx] or "").strip()
                        sku = (row[sku_idx] or "").strip()
                        if vname and sku:
                            sku_by_pid_opt[(pid, re.sub(r"\s+", " ", vname.lower()))] = sku
    except Exception as e:
        print(f"[SKU][WARN] SALES 탭 처리 스킵: {e}")
    
    buckets: Dict[str, Dict[str, List]] = {}
    failures: List[List[str]] = []
    
    def set_if_exists(headers: List[str], row: List[str], name: str, value: str):
        idx = _find_col_index([header_key(h) for h in headers], name)
        if idx >= 0: row[idx] = value
    
    for r in range(media_first - 1, len(media_vals)):
        row = media_vals[r]
        pid = (row[media_hdr.pid] or "").strip() if media_hdr.pid >= 0 and len(row)>media_hdr.pid else ""
        cat = (row[media_hdr.category] or "").strip() if media_hdr.category >= 0 and len(row)>media_hdr.category else ""
        if not pid or not cat: continue
        
        pname = (row[media_hdr.pname] if media_hdr.pname >= 0 else "") or ""
        item_imgs = [(row[i] or "").strip() for i in media_hdr.item_images if i < len(row)]
        
        options = []
        for i, name_col in enumerate(media_hdr.opt_name_cols):
            opt_name = (row[name_col] if name_col < len(row) else "").strip()
            if not opt_name: continue
            img_val = ""
            if i < len(media_hdr.opt_img_cols) and media_hdr.opt_img_cols[i] >= 0:
                img_val = (row[media_hdr.opt_img_cols[i]] if media_hdr.opt_img_cols[i] < len(row) else "").strip()
            options.append((opt_name, img_val))

        top_norm = header_key(top_of_category(cat) or "")
        headers = template_dict.get(top_norm)
        if not headers:
            failures.append([pid, cat, pname, "TEMPLATE_TOPLEVEL_NOT_FOUND", f"top={top_of_category(cat)}"])
            continue

        psku_val = parent_sku_map.get(pid, "")
        
        if not options:
            arr = [""] * len(headers)
            set_if_exists(headers, arr, "category", cat)
            set_if_exists(headers, arr, "product name", pname)
            for k, url in enumerate(item_imgs, start=1):
                if url: set_if_exists(headers, arr, f"item image {k}", url)
            if psku_val: set_if_exists(headers, arr, "parent sku", psku_val)
            b = buckets.setdefault(top_norm, {"headers": headers, "pids": [], "rows": []})
            b["pids"].append([pid]); b["rows"].append(arr)
        else:
            var_label_val = (row[media_hdr.var_label] if media_hdr.var_label >= 0 else "") or "color"
            for (opt_name_raw, opt_img) in options:
                arr = [""] * len(headers)
                set_if_exists(headers, arr, "category", cat)
                set_if_exists(headers, arr, "product name", pname)
                set_if_exists(headers, arr, "variation name1", var_label_val)
                set_if_exists(headers, arr, "option for variation 1", opt_name_raw)
                if opt_img: set_if_exists(headers, arr, "image per variation", opt_img)
                for k, url in enumerate(item_imgs, start=1):
                    if url: set_if_exists(headers, arr, f"item image {k}", url)
                if psku_val: set_if_exists(headers, arr, "parent sku", psku_val)

                opt_key = (pid, re.sub(r"\s+", " ", opt_name_raw.lower()))
                csku_val = sku_by_pid_opt.get(opt_key, "")
                if csku_val:
                    set_if_exists(headers, arr, "sku", csku_val)
                else:
                    failures.append([pid, cat, pname, "SKU_MATCH_NOT_FOUND", f"opt={opt_name_raw}"])
                
                b = buckets.setdefault(top_norm, {"headers": headers, "pids": [], "rows": []})
                b["pids"].append([pid]); b["rows"].append(arr)

    out_matrix: List[List[str]] = []
    for _, pack in buckets.items():
        out_matrix.append([""] + pack["headers"])
        for pid_row, data_row in zip(pack["pids"], pack["rows"]):
            out_matrix.append(pid_row + data_row)

    if out_matrix:
        try:
            tem_ws = safe_worksheet(sh, tem_name)
            with_retry(lambda: tem_ws.clear())
        except Exception:
            tem_ws = with_retry(lambda: sh.add_worksheet(title=tem_name, rows=5000, cols=200))

        max_cols = max(len(r) for r in out_matrix)
        end_a1 = rowcol_to_a1(len(out_matrix), max_cols)
        with_retry(lambda: tem_ws.resize(rows=len(out_matrix) + 10, cols=max_cols + 10))
        with_retry(lambda: tem_ws.update(values=out_matrix, range_name=f"A1:{end_a1}"))
    
    if failures:
        _append_failures(sh, failures)

    print("========== STEP 1 RESULT ==========")
    print(f"TEM 생성 행수: {len(out_matrix) - len(buckets):,}")
    print(f"Failures 기록: {len(failures):,}")
    print("Step 1: Build TEM_OUTPUT Finished.")

# ==============================================================================
# STEP 2: Mandatory 기본값 채우기
# ==============================================================================

def run_step_2(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet):
    """Step 2: TEM_OUTPUT에 Mandatory 기본값 채우기 + 색칠"""
    print("\n[ Automation ] Starting Step 2: Fill Mandatory Defaults...")

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
            cat = (row[c_idx] if c_idx < len(row) else "").strip()
            attr = (row[a_idx] if a_idx < len(row) else "").strip()
            dval = (row[d_idx] if d_idx < len(row) else "").strip()
            if cat and attr:
                out.setdefault((cat or "").strip().lower(), {})[header_key(attr)] = dval
        return out

    sheets = with_retry(lambda: ref.worksheets())
    defaults_map = {}
    for ws in sheets:
        if ws.title.lower().startswith("mandatorydefaults_"):
            for k, d in _read_defaults_ws(ws).items():
                defaults_map.setdefault(k, {}).update(d)

    cat_props_ws = safe_worksheet(ref, get_env("CAT_PROPS_SHEET", "cat props"))
    cat_props_vals = with_retry(lambda: cat_props_ws.get_all_values()) or []
    catprops_map = {}
    if cat_props_vals:
        hdr_keys = [header_key(x) for x in cat_props_vals[0]]
        for r in range(1, len(cat_props_vals)):
            row = cat_props_vals[r]
            cat_raw = (row[0] if len(row) > 0 else "").strip()
            if not cat_raw: continue
            mand_list = [hdr_keys[j] for j, cell in enumerate(row) if str(cell).strip().lower() == "mandatory"]
            if mand_list: catprops_map[(cat_raw or "").strip().lower()] = mand_list

    vals = with_retry(lambda: tem_ws.get_all_values()) or []
    if not vals: print("[!] TEM_OUTPUT 비어 있음."); return

    meta = with_retry(lambda: sh.fetch_sheet_metadata())
    sheet_id = next((s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == tem_name), None)
    if sheet_id is None: print("[!] 시트 ID 찾지 못함."); return

    updates: List[Cell] = []
    color_ranges_by_col = defaultdict(list)
    current_hdr_keys = None
    total_filled = 0

    for r0, row in enumerate(vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_hdr_keys = [header_key(h) for h in row[1:]]
            continue
        if not current_hdr_keys: continue

        pid = (row[0] if len(row) > 0 else "").strip()
        cat_raw = (row[1] if len(row) > 1 else "").strip()
        if not pid or not cat_raw: continue
        norm_cat = (cat_raw or "").strip().lower()

        if norm_cat in catprops_map:
            for attr_norm in catprops_map[norm_cat]:
                j = _find_col_index(current_hdr_keys, attr_norm)
                if j >= 0: color_ranges_by_col[j].append((r0, r0 + 1))
        
        if norm_cat in defaults_map:
            for attr_norm, dval in defaults_map[norm_cat].items():
                if not dval: continue
                j = _find_col_index(current_hdr_keys, attr_norm)
                if j < 0: continue
                col_1based = j + 2
                cur = (row[col_1based - 1] if len(row) >= col_1based else "").strip()
                if not cur or overwrite:
                    updates.append(Cell(row=r0 + 1, col=col_1based, value=dval))
                    total_filled += 1

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="RAW"))

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
            requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": s, "endRowIndex": e, "startColumnIndex": 1 + j, "endColumnIndex": 1 + j + 1}, "cell": {"userEnteredFormat": {"backgroundColor": color}}, "fields": "userEnteredFormat.backgroundColor"}})
    
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
    """
    (개선) Step 3: Reference 시트의 목록을 기준으로, TEM_OUTPUT 행에 고정 FDA 코드를 채웁니다.
    """
    print("\n[ Automation ] Starting Step 3: Fill FDA Code...")
    
    tem_name = get_tem_sheet_name()
    fda_sheet_name = get_env("FDA_CATEGORIES_SHEET_NAME", "TH Cos") # .env에서 시트 이름 읽기
    fda_header = get_env("FDA_HEADER_NAME", "FDA Registration No.")
    FDA_CODE = "10-1-9999999"  # 고정값으로 변경

    try:
        # Reference 시트에서 FDA 대상 카테고리 목록 읽기
        fda_ws = safe_worksheet(ref, fda_sheet_name)
        fda_vals_2d = with_retry(lambda: fda_ws.get_values('A:A', value_render_option='UNFORMATTED_VALUE'))
        fda_vals = [r[0] for r in (fda_vals_2d or []) if r and str(r[0]).strip()]
        # (개선) 전체 경로를 소문자로 변환하여 비교
        target_categories = {str(cat).strip().lower() for cat in fda_vals if str(cat).strip()}
    except Exception as e:
        print(f"[!] '{fda_sheet_name}' 탭을 읽는 데 실패했습니다: {e}. Step 3을 건너<binary data, 2 bytes><binary data, 2 bytes><binary data, 2 bytes>니다.")
        return

    try:
        tem_ws = safe_worksheet(sh, tem_name)
        vals = with_retry(lambda: tem_ws.get_all_values()) or []
    except WorksheetNotFound:
        print(f"[!] {tem_name} 탭 없음. Step1 선행 필요."); return

    if not vals: print("[!] TEM_OUTPUT 비어 있음."); return

    updates: List[Cell] = []
    current_keys, col_category_B, col_fda_B = None, -1, -1
    updated_rows = 0

    for r0, row in enumerate(vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_keys = [header_key(h) for h in row[1:]]
            col_category_B = _find_col_index(current_keys, "category")
            col_fda_B = _find_col_index(current_keys, fda_header)
            continue
        if not current_keys or col_fda_B < 0 or col_category_B < 0: continue

        pid = (row[0] if len(row) > 0 else "").strip()
        if not pid: continue
        
        # (개선) TEM_OUTPUT의 카테고리 값도 전체 경로를 소문자로 변환하여 비교
        category_val_raw = (row[col_category_B + 1] if len(row) > (col_category_B + 1) else "").strip()
        category_val_normalized = category_val_raw.lower()
        
        # (개선) 정규화된 전체 경로가 목록에 있는지 확인
        if category_val_normalized and category_val_normalized in target_categories:
            c_fda_sheet_col = col_fda_B + 2
            cur_fda = (row[c_fda_sheet_col - 1] if len(row) >= c_fda_sheet_col else "").strip()
            
            if not cur_fda or overwrite:
                updates.append(Cell(row=r0 + 1, col=c_fda_sheet_col, value=FDA_CODE))
                updated_rows += 1

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="RAW"))

    print(f"========== STEP 3 FDA RESULT (WRITE) ==========")
    print(f"적용된 셀 수: {updated_rows:,}")
    print("Step 3: Fill FDA Code Finished.")


# ==============================================================================
# STEP 4: 기타 필드 채우기
# ==============================================================================

def run_step_4(sh: gspread.Spreadsheet, ref: gspread.Spreadsheet):
    """Step 4: TEM_OUTPUT의 Stock / Days to ship / Weight / Brand 채우기"""
    print("\n[ Automation ] Starting Step 4: Fill Other Fields...")

    tem_name = get_tem_sheet_name()
    STOCK_VALUE = int(get_env("STEP4_STOCK_VALUE", "1000"))
    DTOS_VALUE  = int(get_env("STEP4_DTOS_VALUE", "1"))

    # (개선) Step 3의 변경 사항이 반영된 최신 데이터를 다시 읽어옴
    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []
    if not tem_vals: print("[!] TEM_OUTPUT 비어 있음."); return

    try:
        margin_ws = safe_worksheet(sh, "MARGIN")
        margin_vals = with_retry(lambda: margin_ws.get_all_values()) or []
    except Exception: margin_vals = []

    try:
        brand_ws = safe_worksheet(ref, "Brand")
        brand_vals = with_retry(lambda: brand_ws.get_all_values()) or []
    except Exception:
        brand_vals = []

    sku_to_weight: Dict[str, str] = {}
    sku_to_brand_name: Dict[str, str] = {}
    if margin_vals:
        mh = margin_vals[0]
        idx_sku = _pick_index_by_candidates(mh, ["sku", "seller_sku"])
        idx_brandn = _pick_index_by_candidates(mh, ["brand", "brand name"])
        idx_wgt = _pick_index_by_candidates(mh, ["weight", "package weight"])
        if idx_sku >= 0:
            for r in range(1, len(margin_vals)):
                row = margin_vals[r]
                sku = (row[idx_sku] if idx_sku < len(row) else "").strip()
                if not sku: continue
                if 0 <= idx_wgt < len(row): sku_to_weight[sku] = (row[idx_wgt] or "").strip()
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
    current_headers = None

    for r in range(len(tem_vals)):
        row = tem_vals[r]
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = row[1:]
            idx_stock_B  = _find_col_index([header_key(h) for h in current_headers], "stock")
            idx_dtos_B   = _find_col_index([header_key(h) for h in current_headers], "daystoship")
            idx_weight_B = _find_col_index([header_key(h) for h in current_headers], "weight")
            idx_brand_B  = _find_col_index([header_key(h) for h in current_headers], "brand")
            idx_sku_B    = _find_col_index([header_key(h) for h in current_headers], "sku")
            continue
        if not current_headers: continue
        
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
            else: failures.append([pid, "", "", "WEIGHT_MAP_MISSING", f"sku={sku_val}"])

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
# STEP 5: Description, Variation Integration, Global SKU Price
# ==============================================================================
def run_step_5(sh: gspread.Spreadsheet):
    """Step 5: Description, Variation Integration, Global SKU Price 채우기"""
    print("\n[ Automation ] Starting Step 5: Fill essential info...")

    tem_name = get_tem_sheet_name()
    # (개선) Step 4의 변경 사항이 반영된 최신 데이터를 다시 읽어옴
    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []

    basic_ws = safe_worksheet(sh, "BASIC")
    basic_vals = with_retry(lambda: basic_ws.get_all_values()) or []
    
    margin_ws = safe_worksheet(sh, "MARGIN")
    margin_vals = with_retry(lambda: margin_ws.get_all_values()) or []

    # --- 데이터 맵 준비 ---
    pid_to_desc = {row[0].strip(): (row[3] if len(row) > 3 else "") for row in basic_vals[1:] if row and row[0].strip()}
    sku_to_price = {row[0].strip(): (row[4] if len(row) > 4 else "") for row in margin_vals[1:] if row and row[0].strip()}
    
    updates: List[Cell] = []
    current_headers = None
    pid_groups = defaultdict(list)
    
    for r_idx, row in enumerate(tem_vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = [header_key(h) for h in row[1:]]
            idx_desc = _find_col_index(current_headers, "productdescription")
            idx_var_integ = _find_col_index(current_headers, "variationintegration")
            idx_price = _find_col_index(current_headers, "globalskuprice")
            idx_sku = _find_col_index(current_headers, "sku")
            continue
        if not current_headers: continue
        
        pid = (row[0] if len(row) > 0 else "").strip()
        if not pid: continue
        
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
            if len(rows) > 1: # Only for variations
                v_code = f"V{pid}"
                for r in rows:
                    updates.append(Cell(row=r, col=idx_var_integ + 2, value=v_code))

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="USER_ENTERED"))

    print("Step 5: Fill essential info Finished.")


# ==============================================================================
# STEP 6: Cover Image URL 생성
# ==============================================================================
def run_step_6(sh: gspread.Spreadsheet, shop_code: str):
    """
    (개선) Step 6: Parent SKU 우선 규칙을 적용하여 Cover image URL을 동적으로 생성합니다.
    """
    print("\n[ Automation ] Starting Step 6: Generate Cover Image URLs...")
    
    tem_name = get_tem_sheet_name()
    # (개선) Step 5의 변경 사항이 반영된 최신 데이터를 다시 읽어옴
    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []
    
    host = get_env("IMAGE_HOSTING_URL", "")
    if not host.endswith("/"): host += "/"

    updates: List[Cell] = []
    current_headers = None

    for r_idx, row in enumerate(tem_vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = [header_key(h) for h in row[1:]]
            idx_cover = _find_col_index(current_headers, "coverimage")
            idx_sku = _find_col_index(current_headers, "sku")
            idx_psku = _find_col_index(current_headers, "parentsku")
            continue
        if not current_headers or idx_cover == -1: continue

        # (개선) Parent SKU와 SKU 값을 모두 가져옵니다.
        psku_val = (row[idx_psku + 1] if idx_psku != -1 and len(row) > idx_psku + 1 else "").strip()
        sku_val = (row[idx_sku + 1] if idx_sku != -1 and len(row) > idx_sku + 1 else "").strip()

        # (개선) URL 생성에 사용할 SKU를 결정합니다 (Parent SKU 우선).
        sku_for_url = psku_val if psku_val else sku_val

        # 사용할 SKU가 있는 경우에만 URL을 생성합니다.
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
    """Step 7: TEM_OUTPUT을 TopLevel Category 단위로 분할하여 엑셀 파일 생성"""
    print("\n[ Automation ] Starting Step 7: Generating final template file...")
    tem_name = get_tem_sheet_name()
    tem_ws = safe_worksheet(sh, tem_name)
    
    all_data = with_retry(lambda: tem_ws.get_values())

    if not all_data:
        print("[!] TEM_OUTPUT sheet is empty. Cannot generate file.")
        return None

    df = pd.DataFrame(all_data)
    
    header_indices = df[df[1].str.lower() == 'category'].index
    
    if header_indices.empty:
        print("[!] No valid header rows found in TEM_OUTPUT.")
        return None
        
    output = BytesIO()
    # 엑셀 생성 엔진을 xlsxwriter로 지정하여 성능 확보
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for i, header_index in enumerate(header_indices):
            start_row = header_index + 1
            end_row = header_indices[i+1] if i + 1 < len(header_indices) else len(df)
            
            chunk_df = df.iloc[start_row:end_row]
            if chunk_df.empty: continue
            
            first_cat = chunk_df.iloc[0, 1] if len(chunk_df.iloc[0]) > 1 else "UNKNOWN"
            top_level_name = top_of_category(first_cat) or "UNKNOWN"
            sheet_name = re.sub(r'[\s/\\*?:\[\]]', '_', top_level_name.title())[:31]

            chunk_df = chunk_df.iloc[:, 1:]

            chunk_df.iloc[:, 0] = chunk_df.iloc[:, 0].str.replace(r'\s*-\s*', '-', regex=True)

            chunk_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    output.seek(0)
    print("Step 7: Final template file generated successfully.")
    return output

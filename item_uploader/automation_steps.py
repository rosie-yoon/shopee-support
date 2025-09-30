# -*- coding: utf-8 -*-
"""
automation_steps.py — consolidated
- Root‑cause fixes for Step 1 not running
- Keeps headers in Step 7 (per request)
- Safer handling of reference sheet, failures, and coloring
"""
from __future__ import annotations

import re
from io import BytesIO
from collections import defaultdict
from typing import Dict, List, Tuple, Optional

import gspread
from gspread.cell import Cell
from gspread.utils import rowcol_to_a1
from gspread.exceptions import WorksheetNotFound
import pandas as pd

from .utils_common import (
    get_env,
    get_bool_env,
    with_retry,
    safe_worksheet,
    header_key,
    top_of_category,
    get_tem_sheet_name,
    hex_to_rgb01,
    open_ref_by_env,
)

# ======================================================================
# helpers
# ======================================================================

def _ensure_ref(ref_obj: Optional[gspread.Spreadsheet]) -> gspread.Spreadsheet:
    """Open the reference spreadsheet from env if not provided."""
    return ref_obj or open_ref_by_env()


def _pick_index_by_candidates(header_row: List[str], candidates: List[str]) -> int:
    """Pick column index by best match among candidate names (exact > contains)."""
    keys = [header_key(x) for x in header_row]
    # exact
    for cand in candidates:
        ck = header_key(cand)
        for i, k in enumerate(keys):
            if k == ck:
                return i
    # contains
    for cand in candidates:
        ck = header_key(cand)
        if not ck:
            continue
        for i, k in enumerate(keys):
            if ck in k:
                return i
    return -1


def _find_col_index(keys: List[str], name: str, extra_alias: List[str] = []) -> int:
    """Find column index from *normalized* header keys list."""
    tgt = header_key(name)
    aliases = [header_key(a) for a in extra_alias] + [tgt]
    # exact
    for i, k in enumerate(keys):
        if k in aliases:
            return i
    # contains
    for i, k in enumerate(keys):
        if any(a and a in k for a in aliases):
            return i
    return -1


def _append_failures(sh: gspread.Spreadsheet, rows: List[List[str]]):
    """Append rows to the 'Failures' tab; auto-resize if needed."""
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
        header = [["PID", "Category", "Name", "Reason", "Detail"]]
        with_retry(lambda: ws.update(values=header + rows, range_name="A1"))


# ======================================================================
# STEP 1: Build TEM_OUTPUT from BASIC + MEDIA (+ SALES mapping)
# ======================================================================

def run_step_1(sh: gspread.Spreadsheet, ref: Optional[gspread.Spreadsheet]):
    print("\n[ Automation ] Starting Step 1: Build TEM_OUTPUT...")

    # Always open the *reference* sheet from env (avoid accidental main==ref)
    ref = open_ref_by_env()

    # sanity: prevent main/ref pointing to the same doc
    sh_id = getattr(sh, "id", None)
    ref_id = getattr(ref, "id", None)
    if sh_id and ref_id and sh_id == ref_id:
        raise RuntimeError("[STEP1] Reference sheet equals the main sheet. Check REFERENCE_* env.")

    basic_header = int(get_env("BASIC_HEADER_ROW", "2"))
    basic_first  = int(get_env("BASIC_FIRST_DATA_ROW", "3"))
    media_header = int(get_env("MEDIA_HEADER_ROW", "2"))
    media_first  = int(get_env("MEDIA_FIRST_DATA_ROW", "6"))
    tem_name     = get_tem_sheet_name()

    # read BASIC/MEDIA
    basic_ws   = safe_worksheet(sh, "BASIC")
    media_ws   = safe_worksheet(sh, "MEDIA")
    basic_vals = with_retry(lambda: basic_ws.get_all_values()) or []
    media_vals = with_retry(lambda: media_ws.get_all_values()) or []

    # ----- 강제 실패: 비어있음/헤더부족은 조용히 return 금지 -----
    if len(basic_vals) < basic_header:
        raise RuntimeError(f"[STEP1] BASIC has fewer than BASIC_HEADER_ROW={basic_header} rows.")
    if len(media_vals) < media_header:
        raise RuntimeError(f"[STEP1] MEDIA has fewer than MEDIA_HEADER_ROW={media_header} rows.")

    # read TemplateDict from reference (env-driven sheet name)
    REF_TEMPLATE_TAB = get_env("TEMPLATE_DICT_SHEET_NAME", "TemplateDict")
    try:
        template_dict_ws = safe_worksheet(ref, REF_TEMPLATE_TAB)
    except Exception as e:
        raise RuntimeError(f"[STEP1] Missing '{REF_TEMPLATE_TAB}' in reference: {e}")
    template_vals = with_retry(lambda: template_dict_ws.get_all_values()) or []
    if len(template_vals) < 2:
        raise RuntimeError("[STEP1] TemplateDict has no valid rows.")
    template_dict = {
        header_key(row[0]): [str(x or "").strip() for x in row[1:]]
        for row in template_vals[1:] if (row[0] or "").strip()
    }

    # parse MEDIA header
    class MediaHeader:
        def __init__(self):
            self.pid = -1
            self.pname = -1
            self.category = -1
            self.cover = -1
            self.item_images: List[int] = []
            self.var_label = -1
            self.opt_name_cols: List[int] = []
            self.opt_img_cols: List[int] = []

    def parse_media_header_row(header_row: List[str]) -> MediaHeader:
        h = MediaHeader()
        keys = [header_key(x) for x in header_row]
        h.pid      = _find_col_index(keys, "productid", ["pid","itemid","ettitleproductid"])
        h.pname    = _find_col_index(keys, "productname", ["itemname","name"])
        h.category = _find_col_index(keys, "category")
        h.cover    = _find_col_index(keys, "coverimage", ["coverimg"])
        h.var_label = _find_col_index(keys, "variationname1", ["variationname","variation"])
        for i, raw in enumerate(header_row):
            if header_key(raw).startswith("itemimage"):
                h.item_images.append(i)
        patt = re.compile(r"^option(\d+)name$")
        for idx, raw in enumerate(header_row):
            m = patt.match(header_key(raw))
            if m:
                n = m.group(1)
                h.opt_name_cols.append(idx)
                img_idx = next((j for j, r in enumerate(header_row)
                                if header_key(r) == f"option{n}image"), -1)
                h.opt_img_cols.append(img_idx)
        return h

    # ----- 헤더 행 인덱스 sanity -----
    header_idx = media_header - 1
    if header_idx < 0 or header_idx >= len(media_vals):
        raise RuntimeError(f"[STEP1] MEDIA_HEADER_ROW={media_header} is out of range.")
    media_hdr = parse_media_header_row(media_vals[header_idx])

    # 필수 컬럼 확인 (없으면 즉시 실패)
    if media_hdr.pid < 0 or media_hdr.category < 0:
        raise RuntimeError("[STEP1] MEDIA header missing required columns: PID or Category.")

    # ----- 데이터 시작 행: 헤더 다음 행과 env 둘 중 큰 값 사용 -----
    start_r = max(media_first - 1, header_idx + 1)

    # SALES mappings (optional)
    parent_sku_map: Dict[str, str] = {}
    sku_by_pid_opt: Dict[Tuple[str, str], str] = {}
    try:
        sales_ws   = safe_worksheet(sh, "SALES")
        sales_vals = with_retry(lambda: sales_ws.get_all_values()) or []
        if sales_vals:
            hdr = sales_vals[0]
            pid_idx     = _pick_index_by_candidates(hdr, ["product id","pid","item id","et_title_product_id"])
            psku_idx    = _pick_index_by_candidates(hdr, ["parent sku","parent_sku","seller sku","seller_sku","et_title_parent_sku"])
            var_name_idx= _pick_index_by_candidates(hdr, ["variation name","option name","option 1 name","variation option","variation","option"])
            sku_idx     = _pick_index_by_candidates(hdr, ["sku","variation sku","child sku","option sku","seller_child_sku","et_title_child_sku"])
            if pid_idx >= 0:
                for r in range(1, len(sales_vals)):
                    row = sales_vals[r]
                    pid = (row[pid_idx] if pid_idx < len(row) else "").strip()
                    if not pid:
                        continue
                    if 0 <= psku_idx < len(row) and (row[psku_idx] or "").strip():
                        parent_sku_map[pid] = row[psku_idx].strip()
                    if 0 <= var_name_idx < len(row) and 0 <= sku_idx < len(row):
                        vname = (row[var_name_idx] or "").strip()
                        sku   = (row[sku_idx] or "").strip()
                        if vname and sku:
                            sku_by_pid_opt[(pid, re.sub(r"\s+", " ", vname.lower()))] = sku
    except Exception as e:
        print(f"[SKU][WARN] skip SALES mapping: {e}")

    # build output rows per TopLevel category
    buckets: Dict[str, Dict[str, List]] = {}
    failures: List[List[str]] = []

    def set_if_exists(headers: List[str], row: List[str], name: str, value: str):
        idx = _find_col_index([header_key(h) for h in headers], name)
        if idx >= 0:
            row[idx] = value

    processed = 0
    for r in range(start_r, len(media_vals)):
        row = media_vals[r]
        pid = (row[media_hdr.pid] or "").strip() if len(row) > media_hdr.pid >= 0 else ""
        cat = (row[media_hdr.category] or "").strip() if len(row) > media_hdr.category >= 0 else ""
        if not pid or not cat:
            continue

        pname = (row[media_hdr.pname] if media_hdr.pname >= 0 and media_hdr.pname < len(row) else "") or ""
        item_imgs = [(row[i] or "").strip() for i in media_hdr.item_images if i < len(row)]
        var_label_val = (row[media_hdr.var_label] or "").strip() if 0 <= media_hdr.var_label < len(row) else ""

        top_norm = (top_of_category(cat) or "").lower()
        headers = template_dict.get(header_key(top_norm), []) or template_dict.get(header_key(cat), [])
        if not headers:
            failures.append([pid, cat, pname, "TEMPL_HEADER_NOT_FOUND", f"top={top_norm}"])
            continue

        psku_val = parent_sku_map.get(pid, "")

        # variations
        options: List[Tuple[str, str]] = []
        if media_hdr.opt_name_cols:
            for name_idx, img_idx in zip(media_hdr.opt_name_cols, media_hdr.opt_img_cols):
                opt_name_raw = (row[name_idx] or "").strip() if 0 <= name_idx < len(row) else ""
                opt_img      = (row[img_idx]  or "").strip() if 0 <= img_idx  < len(row) else ""
                if opt_name_raw:
                    options.append((opt_name_raw, opt_img))
        else:
            options.append(("", ""))  # no variation → single row

        for (opt_name_raw, opt_img) in options:
            arr = [""] * len(headers)
            set_if_exists(headers, arr, "category", cat)
            set_if_exists(headers, arr, "product name", pname)
            set_if_exists(headers, arr, "variation name1", var_label_val)
            set_if_exists(headers, arr, "option for variation 1", opt_name_raw)
            if opt_img:
                set_if_exists(headers, arr, "image per variation", opt_img)
            for k, url in enumerate(item_imgs, start=1):
                if url:
                    set_if_exists(headers, arr, f"item image {k}", url)
            if psku_val:
                set_if_exists(headers, arr, "parent sku", psku_val)

            csku_val = sku_by_pid_opt.get((pid, re.sub(r"\s+", " ", opt_name_raw.lower())))
            if csku_val:
                set_if_exists(headers, arr, "sku", csku_val)
            elif opt_name_raw:
                failures.append([pid, cat, pname, "SKU_MATCH_NOT_FOUND", f"opt={opt_name_raw}"])

            b = buckets.setdefault(top_norm, {"headers": headers, "pids": [], "rows": []})
            b["pids"].append([pid])
            b["rows"].append(arr)
            processed += 1

    # ----- 한 건도 처리 못했으면 명시적 실패 -----
    if processed == 0:
        raise RuntimeError(
            "[STEP1] No MEDIA data rows were processed. "
            "Check MEDIA_HEADER_ROW / MEDIA_FIRST_DATA_ROW or sheet content."
        )

    # flatten to matrix with header rows (A-blank + header in B:)
    out_matrix: List[List[str]] = []
    for _, pack in buckets.items():
        out_matrix.append([""] + pack["headers"])  # header marker row (B: has headers, B1 == 'Category')
        for pid_row, data_row in zip(pack["pids"], pack["rows"]):
            out_matrix.append(pid_row + data_row)

    # write TEM_OUTPUT
    try:
        tem_ws = safe_worksheet(sh, tem_name)
        with_retry(lambda: tem_ws.clear())
    except Exception:
        tem_ws = with_retry(lambda: sh.add_worksheet(title=tem_name, rows=5000, cols=200))
    max_cols = max(len(r) for r in out_matrix) if out_matrix else 2
    end_a1   = rowcol_to_a1(len(out_matrix), max_cols)
    with_retry(lambda: tem_ws.resize(rows=len(out_matrix) + 10, cols=max_cols + 10))
    with_retry(lambda: tem_ws.update(values=out_matrix, range_name=f"A1:{end_a1}"))

    if failures:
        _append_failures(sh, failures)

    print("========== STEP 1 RESULT ==========")
    print(f"TEM rows: {len(out_matrix) - len(buckets):,}")
    print(f"Failures logged: {len(failures):,}")
    print("Step 1: Build TEM_OUTPUT Finished.")

# ======================================================================
# STEP 2: Fill mandatory defaults + color mandatory columns
# ======================================================================

def run_step_2(sh: gspread.Spreadsheet, ref: Optional[gspread.Spreadsheet]):
    print("\n[ Automation ] Starting Step 2: Fill Mandatory Defaults...")
    ref = _ensure_ref(ref)

    tem_name = get_tem_sheet_name()
    color_hex = get_env("COLOR_HEX_MANDATORY", "#FFF9C4")
    overwrite = get_bool_env("OVERWRITE_NONEMPTY", False)

    try:
        tem_ws = safe_worksheet(sh, tem_name)
    except WorksheetNotFound:
        print(f"[!] {tem_name} missing. Run Step 1 first.")
        return

    def _read_defaults_ws(ws) -> Dict[str, Dict[str, str]]:
        vals = with_retry(lambda: ws.get_all_values()) or []
        if not vals:
            return {}
        keys = [header_key(x) for x in vals[0]]
        c_idx = _find_col_index(keys, "category")
        a_idx = _find_col_index(keys, "attribute", ["attr", "property"])
        d_idx = _find_col_index(keys, "defaultvalue", ["default"])
        if min(c_idx, a_idx, d_idx) < 0:
            return {}
        out: Dict[str, Dict[str, str]] = {}
        for r in range(1, len(vals)):
            row = vals[r]
            cat = (row[c_idx] if c_idx < len(row) else "").strip()
            attr = (row[a_idx] if a_idx < len(row) else "").strip()
            dval = (row[d_idx] if d_idx < len(row) else "").strip()
            if cat and attr:
                out.setdefault(cat.strip().lower(), {})[header_key(attr)] = dval
        return out

    # merge all MandatoryDefaults_* in ref
    sheets = with_retry(lambda: ref.worksheets())
    defaults_map: Dict[str, Dict[str, str]] = {}
    for ws in sheets:
        if ws.title.lower().startswith("mandatorydefaults_"):
            for k, d in _read_defaults_ws(ws).items():
                defaults_map.setdefault(k, {}).update(d)

    # read cat props for "Mandatory" marks (column coloring)
    cat_props_ws = safe_worksheet(ref, get_env("CAT_PROPS_SHEET", "cat props"))
    cat_props_vals = with_retry(lambda: cat_props_ws.get_all_values()) or []
    catprops_map: Dict[str, List[str]] = {}
    if cat_props_vals:
        hdr_keys = [header_key(x) for x in cat_props_vals[0]]
        for r in range(1, len(cat_props_vals)):
            row = cat_props_vals[r]
            cat_raw = (row[0] if len(row) > 0 else "").strip()
            if not cat_raw:
                continue
            mand_list = [hdr_keys[j] for j, cell in enumerate(row) if str(cell).strip().lower() == "mandatory"]
            if mand_list:
                catprops_map[cat_raw.strip().lower()] = mand_list

    vals = with_retry(lambda: tem_ws.get_all_values()) or []
    if not vals:
        print("[!] TEM_OUTPUT is empty.")
        return

    # sheetId for batch formatting
    meta = with_retry(lambda: sh.fetch_sheet_metadata())
    sheet_id = next((s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == tem_name), None)
    if sheet_id is None:
        print("[!] sheetId not found; skip coloring.")

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
        norm_cat = cat_raw.strip().lower()

        # coloring
        if norm_cat in catprops_map:
            for attr_norm in catprops_map[norm_cat]:
                j = _find_col_index(current_hdr_keys, attr_norm)
                if j >= 0:
                    color_ranges_by_col[j].append((r0, r0 + 1))

        # fill defaults
        if norm_cat in defaults_map:
            for attr_norm, dval in defaults_map[norm_cat].items():
                if not dval:
                    continue
                j = _find_col_index(current_hdr_keys, attr_norm)
                if j < 0:
                    continue
                col_1based = j + 2  # B is 2
                cur = (row[col_1based - 1] if len(row) >= col_1based else "").strip()
                if not cur or overwrite:
                    updates.append(Cell(row=r0 + 1, col=col_1based, value=dval))
                    total_filled += 1

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="RAW"))

    # merge adjacent color ranges
    def _merge(spans: List[Tuple[int, int]]):
        if not spans:
            return []
        spans.sort()
        merged = [spans[0]]
        for s, e in spans[1:]:
            ls, le = merged[-1]
            if s <= le:
                merged[-1] = (ls, max(le, e))
            else:
                merged.append((s, e))
        return merged

    requests = []
    if sheet_id is not None:
        color = hex_to_rgb01(color_hex)
        for j, spans in color_ranges_by_col.items():
            for s, e in _merge(spans):
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": s,
                            "endRowIndex": e,
                            "startColumnIndex": 1 + j,
                            "endColumnIndex": 1 + j + 1,
                        },
                        "cell": {"userEnteredFormat": {"backgroundColor": color}},
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                })
    if requests:
        with_retry(lambda: sh.batch_update({"requests": requests}))

    print("========== STEP 2 RESULT ==========")
    print(f"filled cells: {total_filled:,}")
    print(f"colored columns: {len(color_ranges_by_col):,}")
    print("Step 2: Fill Mandatory Defaults Finished.")


# ======================================================================
# STEP 3: Fill FDA code for target categories
# ======================================================================

def run_step_3(sh: gspread.Spreadsheet, ref: Optional[gspread.Spreadsheet], overwrite: bool = False):
    print("\n[ Automation ] Starting Step 3: Fill FDA Code...")
    ref = _ensure_ref(ref)

    tem_name = get_tem_sheet_name()
    fda_sheet_name = get_env("FDA_CATEGORIES_SHEET_NAME", "TH Cos")
    fda_header = get_env("FDA_HEADER_NAME", "FDA Registration No.")
    FDA_CODE = "10-1-9999999"

    # read target categories from ref
    try:
        fda_ws = safe_worksheet(ref, fda_sheet_name)
        fda_vals_2d = with_retry(lambda: fda_ws.get_values("A:A", value_render_option="UNFORMATTED_VALUE"))
        fda_vals = [r[0] for r in (fda_vals_2d or []) if r and str(r[0]).strip()]
        target_categories = {str(cat).strip().lower() for cat in fda_vals if str(cat).strip()}
    except Exception as e:
        print(f"[!] cannot read '{fda_sheet_name}': {e} → skip Step 3")
        return

    # read TEM
    try:
        tem_ws = safe_worksheet(sh, tem_name)
        vals = with_retry(lambda: tem_ws.get_all_values()) or []
    except WorksheetNotFound:
        print(f"[!] {tem_name} missing. Run Step 1 first.")
        return
    if not vals:
        print("[!] TEM_OUTPUT is empty.")
        return

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

    print("========== STEP 3 FDA RESULT ==========")
    print(f"updated cells: {updated_rows:,}")
    print("Step 3: Fill FDA Code Finished.")


# ======================================================================
# STEP 4: Fill Stock / Days to ship / Weight / Brand
# ======================================================================

def run_step_4(sh: gspread.Spreadsheet, ref: Optional[gspread.Spreadsheet]):
    print("\n[ Automation ] Starting Step 4: Fill Other Fields...")
    ref = _ensure_ref(ref)

    tem_name = get_tem_sheet_name()
    STOCK_VALUE = int(get_env("STEP4_STOCK_VALUE", "1000"))
    DTOS_VALUE = int(get_env("STEP4_DTOS_VALUE", "1"))

    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []
    if not tem_vals:
        print("[!] TEM_OUTPUT is empty.")
        return

    # helpers
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
        idx_sku = _pick_index_by_candidates(mh, ["sku", "seller_sku"])
        idx_brandn = _pick_index_by_candidates(mh, ["brand", "brand name"])
        idx_wgt = _pick_index_by_candidates(mh, ["weight", "package weight"])
        if idx_sku >= 0:
            for r in range(1, len(margin_vals)):
                row = margin_vals[r]
                sku = (row[idx_sku] if idx_sku < len(row) else "").strip()
                if not sku:
                    continue
                if 0 <= idx_wgt < len(row):
                    sku_to_weight[sku] = (row[idx_wgt] or "").strip()
                if 0 <= idx_brandn < len(row):
                    sku_to_brand_name[sku] = (row[idx_brandn] or "").strip()

    brand_name_to_code: Dict[str, str] = {}
    if brand_vals and len(brand_vals[0]) >= 3:
        for r in range(1, len(brand_vals)):
            row = brand_vals[r]
            if len(row) < 3:
                continue
            bname = (row[1] or "").strip()
            bcode = (row[2] or "").strip()
            if bname:
                brand_name_to_code[re.sub(r"\s+", " ", bname.lower())] = bcode

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
            idx_stock_B = _find_col_index(hdr_keys, "stock")
            idx_dtos_B = _find_col_index(hdr_keys, "daystoship")
            idx_weight_B = _find_col_index(hdr_keys, "weight")
            idx_brand_B = _find_col_index(hdr_keys, "brand")
            idx_sku_B = _find_col_index(hdr_keys, "sku")
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
    print(f"Stock/DTOS/Weight/Brand: {cnt_stock}/{cnt_dtos}/{cnt_weight}/{cnt_brand}")
    print("Step 4: Fill Other Fields Finished.")


# ======================================================================
# STEP 5: Fill description / variation integration / global price
# ======================================================================

def run_step_5(sh: gspread.Spreadsheet):
    print("\n[ Automation ] Starting Step 5: Fill essential info...")

    tem_name = get_tem_sheet_name()
    tem_ws = safe_worksheet(sh, tem_name)
    tem_vals = with_retry(lambda: tem_ws.get_all_values()) or []

    basic_ws = safe_worksheet(sh, "BASIC")
    basic_vals = with_retry(lambda: basic_ws.get_all_values()) or []

    margin_ws = safe_worksheet(sh, "MARGIN")
    margin_vals = with_retry(lambda: margin_ws.get_all_values()) or []

    # maps
    pid_to_desc = {row[0].strip(): (row[3] if len(row) > 3 else "") for row in basic_vals[1:] if row and row[0].strip()}
    sku_to_price = {row[0].strip(): (row[4] if len(row) > 4 else "") for row in margin_vals[1:] if row and row[0].strip()}

    updates: List[Cell] = []
    current_headers: Optional[List[str]] = None
    from collections import defaultdict as _dd
    pid_groups = _dd(list)
    idx_desc = idx_var_integ = idx_price = idx_sku = -1

    for r_idx, row in enumerate(tem_vals):
        if (row[1] if len(row) > 1 else "").strip().lower() == "category":
            current_headers = [header_key(h) for h in row[1:]]
            idx_desc = _find_col_index(current_headers, "productdescription")
            idx_var_integ = _find_col_index(current_headers, "variationintegration")
            idx_price = _find_col_index(current_headers, "globalskuprice")
            idx_sku = _find_col_index(current_headers, "sku")
            continue
        if not current_headers:
            continue

        pid = (row[0] if len(row) > 0 else "").strip()
        if not pid:
            continue
        pid_groups[pid].append(r_idx + 1)

        # 1) description
        if idx_desc != -1:
            desc = pid_to_desc.get(pid, "")
            updates.append(Cell(row=r_idx + 1, col=idx_desc + 2, value=desc))

        # 3) price
        if idx_price != -1 and idx_sku != -1:
            sku_val = (row[idx_sku + 1] if len(row) > idx_sku + 1 else "").strip()
            if sku_val:
                price = sku_to_price.get(sku_val, "")
                updates.append(Cell(row=r_idx + 1, col=idx_price + 2, value=price))

    # 2) variation integration code
    if idx_var_integ != -1:
        for pid, rows in pid_groups.items():
            if len(rows) > 1:  # variations only
                v_code = f"V{pid}"
                for r in rows:
                    updates.append(Cell(row=r, col=idx_var_integ + 2, value=v_code))

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="USER_ENTERED"))

    print("Step 5: Fill essential info Finished.")


# ======================================================================
# STEP 6: Generate cover image URLs (Parent SKU preferred)
# ======================================================================

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
            idx_sku = _find_col_index(current_headers, "sku")
            idx_psku = _find_col_index(current_headers, "parentsku")
            continue
        if not current_headers or idx_cover == -1:
            continue

        psku_val = (row[idx_psku + 1] if idx_psku != -1 and len(row) > idx_psku + 1 else "").strip()
        sku_val = (row[idx_sku + 1] if idx_sku != -1 and len(row) > idx_sku + 1 else "").strip()
        sku_for_url = psku_val if psku_val else sku_val
        if sku_for_url:
            url = f"{host}{sku_for_url}_C_{shop_code}.jpg"
            updates.append(Cell(row=r_idx + 1, col=idx_cover + 2, value=url))

    if updates:
        with_retry(lambda: tem_ws.update_cells(updates, value_input_option="USER_ENTERED"))

    print("Step 6: Generate Cover Image URLs Finished.")


# ======================================================================
# STEP 7: Split by top-level category & build final Excel (headers kept)
# ======================================================================

def run_step_7(sh: gspread.Spreadsheet) -> Optional[BytesIO]:
    print("\n[ Automation ] Starting Step 7: Generating final template file...")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # 0) Failure tab first — try multiple known sources
        failure_sources = ["FAILURE_OUTPUT", "Failures"]
        wrote_failure = False
        for failure_sheet_name in failure_sources:
            try:
                failure_ws = safe_worksheet(sh, failure_sheet_name)
                failure_data = with_retry(lambda: failure_ws.get_all_records())
                if failure_data:
                    failure_df = pd.DataFrame(failure_data)
                    failure_df.to_excel(writer, sheet_name="Failure", index=False, header=True)
                    wrote_failure = True
                    break
            except Exception:
                pass
        if not wrote_failure:
            # write empty Failure tab to keep structure predictable
            pd.DataFrame([{"Info": "No failures"}]).to_excel(
                writer, sheet_name="Failure", index=False, header=True
            )

        # 1) Split TEM_OUTPUT
        tem_name = get_tem_sheet_name()
        tem_ws = safe_worksheet(sh, tem_name)
        all_data = with_retry(lambda: tem_ws.get_all_values())
        if not all_data or len(all_data) < 2:
            print("[!] TEM_OUTPUT has no data.")
            output.seek(0)
            return output if output.getbuffer().nbytes > 0 else None

        df = pd.DataFrame(all_data)
        col1 = df.iloc[:, 1].astype(str).str.lower()
        header_indices = col1.index[col1.eq("category")]
        if len(header_indices) == 0:
            print("[!] No header rows found in TEM_OUTPUT.")
            output.seek(0)
            return output if output.getbuffer().nbytes > 0 else None

        for i, header_idx in enumerate(header_indices):
            start_row = header_idx
            end_row = header_indices[i + 1] if i + 1 < len(header_indices) else len(df)
            chunk_with_header_df = df.iloc[start_row:end_row]
            if len(chunk_with_header_df) < 2:
                continue

            # keep header per sheet (request)
            new_header = chunk_with_header_df.iloc[0]
            chunk_data_df = chunk_with_header_df[1:].copy()
            chunk_data_df.columns = new_header

            try:
                first_cat = str(chunk_data_df.iloc[0, 1] or "")
            except Exception:
                first_cat = "UNKNOWN"
            top_level_name = top_of_category(first_cat) or "UNKNOWN"
            sheet_name = re.sub(r"[\s/\\*?:\[\]]", "_", top_level_name.title())[:31] or "Sheet1"

            # drop column A (PID) from Excel export; keep columns starting from B
            sub_df = chunk_data_df.iloc[:, 1:].copy()
            try:
                sub_df.iloc[:, 0] = sub_df.iloc[:, 0].astype(str).str.replace(r"\s*-\s*", "-", regex=True)
            except Exception:
                pass

            sub_df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

    output.seek(0)
    print("✅ Step 7: Final template file generated successfully.")
    return output

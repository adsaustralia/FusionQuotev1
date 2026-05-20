import io
import json
import re
import traceback
import zipfile
import xml.etree.ElementTree as ET

import streamlit as st

APP_VERSION = "V2.3 Minimal Stable"


def norm(v):
    return "" if v is None else str(v).strip()


def safe_col(v, default):
    v = (v or "").strip().upper()
    return v if re.fullmatch(r"[A-Z]{1,3}", v) else default


def quote_sheet(name):
    return str(name).replace("'", "''")


def get_sheet_names_fast(xlsx_bytes):
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
        root = ET.fromstring(z.read("xl/workbook.xml"))
    return [el.attrib.get("name", "") for el in root.findall("main:sheets/main:sheet", ns) if el.attrib.get("name")]


def default_sheet(sheets, wanted, fallback=0):
    for s in sheets:
        if s.lower().strip() == wanted.lower().strip():
            return sheets.index(s)
    return min(fallback, len(sheets) - 1) if sheets else 0


def load_openpyxl_bits():
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import column_index_from_string, get_column_letter
    return load_workbook, Font, PatternFill, column_index_from_string, get_column_letter


def detect_multiplier(text):
    t = norm(text).upper()
    m = re.search(r"\bSET\s*(?:OF)?\s*(\d{1,3})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 500:
            return n, "RED", f"set of {n}"
    m = re.search(r"\b(?:1\s*)?PACK\s*=\s*(\d{1,5})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 10000:
            return n, "RED", f"pack = {n}"
    m = re.search(r"\bPACK\s+OF\s+(\d{1,5})\b", t)
    if m:
        n = int(m.group(1))
        if 1 < n <= 10000:
            return n, "RED", f"pack of {n}"
    if "PACK" in t or "SET" in t:
        return 1, "ORANGE", "unclear pack/set wording"
    return 1, "", ""


def size_to_sqm(size_text):
    s = norm(size_text).lower().replace("×", "x").replace(",", "")
    nums = re.findall(r"(\d+(?:\.\d+)?)", s)
    if len(nums) >= 2:
        w, h = float(nums[0]), float(nums[1])
        if "cm" in s and "mm" not in s:
            return (w / 100.0) * (h / 100.0)
        if re.search(r"\bm\b", s) and "mm" not in s and "cm" not in s:
            return w * h
        return (w / 1000.0) * (h / 1000.0)
    if len(nums) == 1 and ("dia" in s or "ø" in s or "round" in s):
        d = float(nums[0])
        if "cm" in s and "mm" not in s:
            dm = d / 100.0
        elif re.search(r"\bm\b", s) and "mm" not in s and "cm" not in s:
            dm = d
        else:
            dm = d / 1000.0
        return 3.141592653589793 * (dm / 2.0) ** 2
    return 0.0


def ignored_array(ignore_list):
    vals = [x.strip().upper().replace('"', '') for x in ignore_list if x.strip()]
    if not vals:
        vals = ["NZ"]
    return "{" + ",".join('"' + x + '"' for x in vals) + "}"


def clean_qty_formula(col, total_row, scan_start, scan_end, country_col, ignored, multiplier):
    arr = ignored_array(ignored)
    base = f"({col}${total_row}-SUMPRODUCT(({col}${scan_start}:{col}${scan_end})*(--ISNUMBER(MATCH(UPPER(${country_col}${scan_start}:${country_col}${scan_end}),{arr},0)))))"
    return f"={base}*{multiplier}" if multiplier != 1 else f"={base}"


def ds_lookup_formula(col, cfg):
    rs = quote_sheet(cfg["reference_sheet"])
    return (
        f"=IFERROR(INDEX('{rs}'!${cfg['ref_ds_col']}${cfg['ref_start_row']}:${cfg['ref_ds_col']}${cfg['ref_end_row']},"
        f"MATCH(1,INDEX(('{rs}'!${cfg['ref_name_col']}${cfg['ref_start_row']}:${cfg['ref_name_col']}${cfg['ref_end_row']}={col}${cfg['name_row']})*"
        f"('{rs}'!${cfg['ref_size_col']}${cfg['ref_start_row']}:${cfg['ref_size_col']}${cfg['ref_end_row']}={col}${cfg['size_row']}),0),0)),\"\")"
    )


def get_stock_options(xlsx_bytes, cfg):
    load_workbook, _, _, col_index, _ = load_openpyxl_bits()
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    try:
        ws = wb[cfg["working_sheet"]]
        start_i = col_index(cfg["item_start_col"])
        end_i = col_index(cfg["item_end_col"])
        result, seen = [], set()
        for c in range(start_i, end_i + 1):
            v = norm(ws.cell(cfg["stock_row"], c).value)
            key = v.upper()
            if v and key not in seen:
                seen.add(key)
                result.append(v)
        return result
    finally:
        wb.close()


def detect_country_scan_end(xlsx_bytes, cfg):
    load_workbook, _, _, col_index, _ = load_openpyxl_bits()
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    try:
        ws = wb[cfg["working_sheet"]]
        c = col_index(cfg["country_col"])
        start = cfg["total_qty_row"] + 1
        last = start
        tokens = {"NZ", "AUS", "AU", "AUSTRALIA", "NEW ZEALAND", "FIJI", "SG", "SINGAPORE"}
        for r in range(start, min(ws.max_row, 500) + 1):
            if norm(ws.cell(r, c).value).upper() in tokens:
                last = r
        return max(last, start)
    finally:
        wb.close()


def build_workbook(xlsx_bytes, cfg, rates):
    load_workbook, Font, PatternFill, col_index, get_col = load_openpyxl_bits()
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    wbv = load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    try:
        ws = wb[cfg["working_sheet"]]
        wsv = wbv[cfg["working_sheet"]]
        start_i, end_i = col_index(cfg["item_start_col"]), col_index(cfg["item_end_col"])
        selected = set(cfg.get("selected_stocks", []))
        ds_factor = 1 + float(cfg.get("ds_loading", 20.0)) / 100.0
        red = PatternFill("solid", fgColor="FFC7CE")
        orange = PatternFill("solid", fgColor="FCE4D6")
        yellow = PatternFill("solid", fgColor="FFF2CC")
        blue = PatternFill("solid", fgColor="1F4E78")
        white = Font(color="FFFFFF", bold=True)
        def head(cell):
            cell.fill = blue
            cell.font = white
        for row, label in [(cfg["ds_row"], "DS/SS"), (cfg["clean_qty_row"], "Clean Qty"), (cfg["multiplier_row"], "Multiplier"), (cfg["sqm_row"], "SQM"), (cfg["price_row"], "Price")]:
            ws.cell(row, 1).value = label
            head(ws.cell(row, 1))
        audit = []
        for c in range(start_i, end_i + 1):
            col = get_col(c)
            name = norm(wsv.cell(cfg["name_row"], c).value or ws.cell(cfg["name_row"], c).value)
            size = norm(wsv.cell(cfg["size_row"], c).value or ws.cell(cfg["size_row"], c).value)
            stock = norm(wsv.cell(cfg["stock_row"], c).value or ws.cell(cfg["stock_row"], c).value)
            multiplier, flag, reason = detect_multiplier(name)
            ws.cell(cfg["ds_row"], c).value = ds_lookup_formula(col, cfg)
            ws.cell(cfg["clean_qty_row"], c).value = clean_qty_formula(col, cfg["total_qty_row"], cfg["country_scan_start"], cfg["country_scan_end"], cfg["country_col"], cfg["ignored_countries"], multiplier)
            ws.cell(cfg["multiplier_row"], c).value = multiplier
            sqm_each = size_to_sqm(size)
            ws.cell(cfg["sqm_row"], c).value = f"={col}${cfg['clean_qty_row']}*{sqm_each:.6f}" if sqm_each else ""
            rate = float(rates.get(stock, 0) or 0)
            if stock in selected and rate > 0:
                ws.cell(cfg["price_row"], c).value = f'=IF(UPPER({col}${cfg["ds_row"]})="DS",{col}${cfg["sqm_row"]}*{rate}*{ds_factor},{col}${cfg["sqm_row"]}*{rate})'
            if flag == "RED":
                ws.cell(cfg["name_row"], c).fill = red
                ws.cell(cfg["clean_qty_row"], c).fill = red
                audit.append([col, name, size, stock, multiplier, flag, reason])
            elif flag == "ORANGE":
                ws.cell(cfg["name_row"], c).fill = orange
                audit.append([col, name, size, stock, multiplier, flag, reason])
            if stock in selected:
                ws.cell(cfg["stock_row"], c).fill = yellow
        for sheet_name in ["Qty Multiplier Audit", "Stock SQM Summary"]:
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
        aud = wb.create_sheet("Qty Multiplier Audit")
        aud.append(["Column", "Name", "Size", "Stock", "Multiplier", "Flag", "Reason"])
        for cell in aud[1]:
            head(cell)
        for row in audit:
            aud.append(row)
        sm = wb.create_sheet("Stock SQM Summary")
        sm.append(["Stock", "Rate / SQM", "Total SQM", "Estimated Price", "DS Loading %"])
        for cell in sm[1]:
            head(cell)
        work = quote_sheet(cfg["working_sheet"])
        for r, stock in enumerate(cfg.get("selected_stocks", []), start=2):
            sm.cell(r, 1).value = stock
            sm.cell(r, 2).value = float(rates.get(stock, 0) or 0)
            sm.cell(r, 3).value = f'=SUMIF(\'{work}\'!${cfg["item_start_col"]}${cfg["stock_row"]}:${cfg["item_end_col"]}${cfg["stock_row"]},A{r},\'{work}\'!${cfg["item_start_col"]}${cfg["sqm_row"]}:${cfg["item_end_col"]}${cfg["sqm_row"]})'
            sm.cell(r, 4).value = f'=SUMIF(\'{work}\'!${cfg["item_start_col"]}${cfg["stock_row"]}:${cfg["item_end_col"]}${cfg["stock_row"]},A{r},\'{work}\'!${cfg["item_start_col"]}${cfg["price_row"]}:${cfg["item_end_col"]}${cfg["price_row"]})'
            sm.cell(r, 5).value = float(cfg.get("ds_loading", 20.0))
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()
    finally:
        try:
            wbv.close()
            wb.close()
        except Exception:
            pass


def main():
    st.set_page_config(page_title="Excel Formula Fusion", layout="wide")
    st.title("Excel Formula Fusion")
    st.caption(APP_VERSION)
    st.warning("Stable recovery version: no workbook reading happens until you press a button.")

    defaults = {"uploaded_bytes": None, "uploaded_name": "", "sheets": [], "cfg": None, "rates": {}}
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx", "xlsm"])
    if uploaded is not None:
        st.session_state.uploaded_bytes = uploaded.getvalue()
        st.session_state.uploaded_name = uploaded.name
        st.success(f"Uploaded: {uploaded.name}")

    if not st.session_state.uploaded_bytes:
        st.info("Upload an Excel file to start.")
        return

    st.write(f"Current file: **{st.session_state.uploaded_name}**")
    if st.button("Read workbook / detect sheets"):
        try:
            st.session_state.sheets = get_sheet_names_fast(st.session_state.uploaded_bytes)
            st.success("Sheets detected: " + ", ".join(st.session_state.sheets))
        except Exception:
            st.error("Could not read workbook sheets")
            st.code(traceback.format_exc())
            return

    sheets = st.session_state.sheets
    if not sheets:
        st.info("Press **Read workbook / detect sheets** before continuing.")
        return

    with st.form("mapping"):
        st.subheader("Mapping")
        a, b = st.columns(2)
        with a:
            working_sheet = st.selectbox("Working sheet", sheets, index=default_sheet(sheets, "DL ANZ ALLOCATION"))
            name_row = int(st.number_input("Name row", min_value=1, value=4))
            size_row = int(st.number_input("Size row", min_value=1, value=5))
            stock_row = int(st.number_input("Stock row", min_value=1, value=6))
            total_qty_row = int(st.number_input("Original total qty row", min_value=1, value=7))
            country_col = safe_col(st.text_input("Country column", value="I"), "I")
            ignored = st.text_input("Ignore countries", value="NZ")
        with b:
            reference_sheet = st.selectbox("Reference sheet", sheets, index=default_sheet(sheets, "PRINT DB", 1 if len(sheets) > 1 else 0))
            ref_name_col = safe_col(st.text_input("Reference name column", value="C"), "C")
            ref_size_col = safe_col(st.text_input("Reference size column", value="E"), "E")
            ref_ds_col = safe_col(st.text_input("Reference DS/SS column", value="F"), "F")
            ref_start_row = int(st.number_input("Reference start row", min_value=1, value=12))
            ref_end_row = int(st.number_input("Reference end row", min_value=1, value=141))
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            item_start_col = safe_col(st.text_input("Start column", value="AC"), "AC")
            item_end_col = safe_col(st.text_input("End column", value="IG"), "IG")
        with c2:
            ds_row = int(st.number_input("DS/SS row", min_value=1, value=168))
            clean_qty_row = int(st.number_input("Clean qty row", min_value=1, value=169))
        with c3:
            multiplier_row = int(st.number_input("Multiplier row", min_value=1, value=170))
            sqm_row = int(st.number_input("SQM row", min_value=1, value=171))
        with c4:
            price_row = int(st.number_input("Price row", min_value=1, value=172))
            ds_loading = float(st.number_input("DS loading %", min_value=0.0, value=20.0, step=1.0))
        submitted = st.form_submit_button("Apply mapping / refresh stock list")

    if submitted:
        try:
            cfg = {
                "working_sheet": working_sheet, "reference_sheet": reference_sheet,
                "name_row": name_row, "size_row": size_row, "stock_row": stock_row, "total_qty_row": total_qty_row,
                "country_col": country_col, "ignored_countries": [x.strip().upper() for x in ignored.split(",") if x.strip()],
                "ref_name_col": ref_name_col, "ref_size_col": ref_size_col, "ref_ds_col": ref_ds_col,
                "ref_start_row": ref_start_row, "ref_end_row": ref_end_row,
                "item_start_col": item_start_col, "item_end_col": item_end_col,
                "ds_row": ds_row, "clean_qty_row": clean_qty_row, "multiplier_row": multiplier_row, "sqm_row": sqm_row, "price_row": price_row,
                "ds_loading": ds_loading,
            }
            cfg["country_scan_start"] = total_qty_row + 1
            cfg["country_scan_end"] = detect_country_scan_end(st.session_state.uploaded_bytes, cfg)
            cfg["stock_options"] = get_stock_options(st.session_state.uploaded_bytes, cfg)
            cfg["selected_stocks"] = []
            st.session_state.cfg = cfg
            st.success(f"Mapping saved. Stock names found: {len(cfg['stock_options'])}. Country scan rows: {cfg['country_scan_start']}:{cfg['country_scan_end']}")
        except Exception:
            st.error("Mapping failed")
            st.code(traceback.format_exc())

    cfg = st.session_state.cfg
    if not cfg:
        return

    st.subheader("Stock/material rates")
    selected = st.multiselect("Pick stock/material to calculate SQM and rate", cfg.get("stock_options", []), default=cfg.get("selected_stocks", []))
    cfg["selected_stocks"] = selected
    if selected:
        with st.form("rates"):
            staged = {}
            for stock in selected:
                staged[stock] = st.number_input(f"Rate $/sqm — {stock}", min_value=0.0, value=float(st.session_state.rates.get(stock, 0.0)), step=0.1, format="%.2f")
            if st.form_submit_button("Refresh / Update Rates"):
                st.session_state.rates.update(staged)
                st.success("Rates updated.")
    st.download_button("Download stock rate memory JSON", json.dumps(st.session_state.rates, indent=2).encode("utf-8"), "stock_rate_memory.json", "application/json")

    st.subheader("Generate")
    if st.button("Generate Excel Workbook"):
        try:
            with st.spinner("Generating workbook..."):
                output = build_workbook(st.session_state.uploaded_bytes, cfg, st.session_state.rates)
            st.success("Workbook generated.")
            st.download_button("Download Excel Workbook", output, "formula_fusion_" + st.session_state.uploaded_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception:
            st.error("Generation failed")
            st.code(traceback.format_exc())


try:
    main()
except Exception:
    try:
        st.error("Startup error")
        st.code(traceback.format_exc())
    except Exception:
        raise

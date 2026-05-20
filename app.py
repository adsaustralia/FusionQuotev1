import io
import json
import re
import traceback
from copy import copy
from typing import Dict, List, Tuple, Optional

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

st.set_page_config(page_title="Excel Formula Fusion", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"] { color: #111111 !important; }
.stApp { background: #f7f7f7; }
section[data-testid="stSidebar"] { background: #ffffff !important; }
input, textarea, select { color: #111111 !important; background-color: #ffffff !important; }
div[data-baseweb="select"] * { color: #111111 !important; }
div[data-baseweb="input"] * { color: #111111 !important; }
button, .stDownloadButton button { color: #111111 !important; background-color: #ffffff !important; border: 1px solid #999 !important; }
.primary-box { background: #ffffff; padding: 1rem; border-radius: 0.7rem; border: 1px solid #ddd; }
.warn-box { background: #fff3cd; padding: 0.8rem; border-radius: 0.5rem; border: 1px solid #ffda6a; }
.good-box { background: #d1e7dd; padding: 0.8rem; border-radius: 0.5rem; border: 1px solid #badbcc; }
</style>
""", unsafe_allow_html=True)

RED_FILL = PatternFill("solid", fgColor="FFC7CE")
ORANGE_FILL = PatternFill("solid", fgColor="FCE4D6")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
GREEN_FILL = PatternFill("solid", fgColor="D9EAD3")


def safe_text(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", safe_text(s)).strip().upper()


def parse_col(value: str, default: str) -> str:
    value = safe_text(value).upper()
    if not value:
        return default
    value = re.sub(r"[^A-Z]", "", value)
    return value or default


def copy_cell_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def detect_multiplier(name: str) -> Tuple[int, str, str]:
    """Return multiplier, status, reason. status: exact/doubt/none."""
    txt = norm(name)
    if not txt:
        return 1, "none", ""
    m = re.search(r"\bSET\s+OF\s+(\d+)\b", txt)
    if m:
        return int(m.group(1)), "exact", f"set of {m.group(1)}"
    m = re.search(r"\b1\s*PACK\s*=\s*(\d+)\b", txt)
    if m:
        return int(m.group(1)), "exact", f"1 pack = {m.group(1)}"
    m = re.search(r"\bPACK\s+OF\s+(\d+)\b", txt)
    if m:
        return int(m.group(1)), "exact", f"pack of {m.group(1)}"
    if "SET" in txt or "PACK" in txt:
        return 1, "doubt", "contains set/pack but no safe multiplier"
    return 1, "none", ""


def parse_size_to_sqm(size_text: str) -> Optional[float]:
    txt = safe_text(size_text).lower().replace("×", "x")
    nums = re.findall(r"(\d+(?:\.\d+)?)", txt)
    if len(nums) < 2:
        return None
    w = float(nums[0])
    h = float(nums[1])
    # default mm; cm if explicitly contains cm and not mm
    if "cm" in txt and "mm" not in txt:
        return (w / 100) * (h / 100)
    return (w / 1000) * (h / 1000)


def get_sheet_names(file_bytes: bytes) -> List[str]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=False)
    names = wb.sheetnames
    wb.close()
    return names


def get_values_for_stock_options(file_bytes: bytes, sheet_name: str, stock_row: int, start_col: str, end_col: str) -> List[str]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb[sheet_name]
    sc = column_index_from_string(start_col)
    ec = column_index_from_string(end_col)
    stocks = []
    for c in range(sc, ec + 1):
        val = safe_text(ws.cell(stock_row, c).value)
        if val and val not in stocks:
            stocks.append(val)
    wb.close()
    return stocks


def build_workbook(
    file_bytes: bytes,
    working_sheet: str,
    reference_sheet: str,
    name_row: int,
    size_row: int,
    stock_row: int,
    total_qty_row: int,
    country_col: str,
    start_col: str,
    end_col: str,
    ignore_countries: List[str],
    ds_output_row: int,
    clean_qty_output_row: int,
    sqm_output_row: int,
    price_output_row: int,
    ref_name_col: str,
    ref_size_col: str,
    ref_ds_col: str,
    ref_stock_col: str,
    ref_start_row: int,
    ref_end_row: int,
    selected_stocks: List[str],
    stock_rates: Dict[str, float],
    ds_loading_pct: float,
) -> bytes:
    """Generate workbook with central stock-rate sheet.

    Important design rule:
    - Streamlit writes formulas and a STOCK RATES sheet.
    - Excel remains the calculation engine.
    - If the user changes one stock rate in STOCK RATES, all formulas update automatically.
    - DS loading is applied only when the DS/SS output row for that item equals DS.
    """
    wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
    wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb[working_sheet]
    ws_values = wb_values[working_sheet]

    sc = column_index_from_string(start_col)
    ec = column_index_from_string(end_col)
    first_col = get_column_letter(sc)
    last_col = get_column_letter(ec)
    country_col = parse_col(country_col, "I")
    ignored = [c.strip().upper() for c in ignore_countries if c.strip()]
    ds_factor = 1 + (float(ds_loading_pct) / 100.0)

    # Replace generated sheets to avoid stale formulas.
    generated_sheets = ["STOCK RATES", "Stock SQM Summary", "Qty Multiplier Audit"]
    for sname in generated_sheets:
        if sname in wb.sheetnames:
            del wb[sname]

    rate_ws = wb.create_sheet("STOCK RATES")
    sum_ws = wb.create_sheet("Stock SQM Summary")
    audit_ws = wb.create_sheet("Qty Multiplier Audit")

    # Central rate table. This is the key feature.
    rate_ws.append(["Stock/Material", "Rate per SQM"])
    rate_stocks = list(dict.fromkeys(list(selected_stocks) + list(stock_rates.keys())))
    for stock in rate_stocks:
        rate_ws.append([stock, float(stock_rates.get(stock, 0) or 0)])
        rate_ws.cell(rate_ws.max_row, 2).number_format = '$#,##0.00'

    # Output labels
    label_col = max(1, sc - 1)
    ws.cell(ds_output_row, label_col).value = "DS/SS Lookup"
    ws.cell(clean_qty_output_row, label_col).value = "Clean Qty"
    ws.cell(sqm_output_row, label_col).value = "SQM"
    ws.cell(price_output_row, label_col).value = "Price"

    # We subtract ignored-country quantities from total row.
    # Start from row below total qty and stop before generated output rows.
    scan_start_row = total_qty_row + 1
    scan_end_row = min(ws.max_row, min(ds_output_row, clean_qty_output_row, sqm_output_row, price_output_row) - 1)
    if scan_end_row < scan_start_row:
        scan_end_row = ws.max_row

    # Main generated formulas
    for c in range(sc, ec + 1):
        col = get_column_letter(c)
        name_val = safe_text(ws_values.cell(name_row, c).value) or safe_text(ws.cell(name_row, c).value)
        size_val = safe_text(ws_values.cell(size_row, c).value) or safe_text(ws.cell(size_row, c).value)
        sqm_each = parse_size_to_sqm(size_val) or 0.0
        multiplier, status, reason = detect_multiplier(name_val)

        for out_r in [ds_output_row, clean_qty_output_row, sqm_output_row, price_output_row]:
            copy_cell_style(ws.cell(total_qty_row, c), ws.cell(out_r, c))

        ws.cell(ds_output_row, c).value = (
            f'=IFERROR(INDEX(\'{reference_sheet}\'!${ref_ds_col}${ref_start_row}:${ref_ds_col}${ref_end_row},'
            f'MATCH(1,INDEX((\'{reference_sheet}\'!${ref_name_col}${ref_start_row}:${ref_name_col}${ref_end_row}={col}${name_row})*'
            f'(\'{reference_sheet}\'!${ref_size_col}${ref_start_row}:${ref_size_col}${ref_end_row}={col}${size_row}),0),0)),"")'
        )

        ignored_terms = []
        for country in ignored:
            ignored_terms.append(
                f'SUMIF(${country_col}${scan_start_row}:${country_col}${scan_end_row},"{country}",{col}${scan_start_row}:{col}${scan_end_row})'
            )
        ignored_expr = "+".join(ignored_terms) if ignored_terms else "0"
        ws.cell(clean_qty_output_row, c).value = f'=({col}${total_qty_row}-({ignored_expr}))*{multiplier}'
        ws.cell(sqm_output_row, c).value = f'={col}${clean_qty_output_row}*{sqm_each}'

        # Centralised stock rate lookup. Change a rate in STOCK RATES and all prices update.
        rate_lookup = f"VLOOKUP({col}${stock_row},'STOCK RATES'!$A:$B,2,FALSE)"
        ws.cell(price_output_row, c).value = (
            f'=IFERROR(IF(UPPER({col}${ds_output_row})="DS",'
            f'{col}${sqm_output_row}*{rate_lookup}*{ds_factor},'
            f'{col}${sqm_output_row}*{rate_lookup}),0)'
        )
        ws.cell(price_output_row, c).number_format = '$#,##0.00'

        if status == "exact":
            for r in [name_row, clean_qty_output_row, sqm_output_row, price_output_row]:
                ws.cell(r, c).fill = RED_FILL
            audit_ws.append([col, name_val, multiplier, "MULTIPLIED", reason])
        elif status == "doubt":
            for r in [name_row, clean_qty_output_row]:
                ws.cell(r, c).fill = ORANGE_FILL
            audit_ws.append([col, name_val, multiplier, "CHECK - NOT MULTIPLIED", reason])

    # Summary sheet is formula-driven and references the same STOCK RATES table.
    sum_ws.append(["Stock/Material", "Rate", "Total SQM", "SS/Other SQM", "DS SQM", "DS Loading %", "Estimated Value"])
    sum_ws.append(["NOTE", "", "", "", "", "", "Change rates in STOCK RATES column B; all formulas update automatically in Excel."])

    for stock in selected_stocks:
        next_row = sum_ws.max_row + 1
        total_sqm_formula = (
            f"=SUMIF('{working_sheet}'!${first_col}${stock_row}:${last_col}${stock_row},"
            f"A{next_row},'{working_sheet}'!${first_col}${sqm_output_row}:${last_col}${sqm_output_row})"
        )
        ds_sqm_formula = (
            f'=SUMPRODUCT((\'{working_sheet}\'!${first_col}${stock_row}:${last_col}${stock_row}=A{next_row})*'
            f'(UPPER(\'{working_sheet}\'!${first_col}${ds_output_row}:${last_col}${ds_output_row})="DS")*'
            f'(\'{working_sheet}\'!${first_col}${sqm_output_row}:${last_col}${sqm_output_row}))'
        )
        ss_sqm_formula = f"=C{next_row}-E{next_row}"
        rate_formula = f"=IFERROR(VLOOKUP(A{next_row},'STOCK RATES'!$A:$B,2,FALSE),0)"
        estimated_value_formula = f"=D{next_row}*B{next_row}+E{next_row}*B{next_row}*(1+F{next_row}/100)"
        sum_ws.append([stock, rate_formula, total_sqm_formula, ss_sqm_formula, ds_sqm_formula, ds_loading_pct, estimated_value_formula])
        sum_ws.cell(next_row, 2).number_format = '$#,##0.00'
        sum_ws.cell(next_row, 7).number_format = '$#,##0.00'

    for ws2 in [rate_ws, sum_ws, audit_ws]:
        for cell in ws2[1]:
            cell.fill = GREEN_FILL
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for col_cells in ws2.columns:
            max_len = max(len(safe_text(cell.value)) for cell in col_cells)
            ws2.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 70)

    output = io.BytesIO()
    wb.save(output)
    wb.close()
    wb_values.close()
    output.seek(0)
    return output.getvalue()

def main():
    st.title("Excel Formula Fusion")
    st.caption("Install-safe build: DS loading applies only when DS/SS row = DS")

    if "file_bytes" not in st.session_state:
        st.session_state.file_bytes = None
    if "sheet_names" not in st.session_state:
        st.session_state.sheet_names = []
    if "stock_rates" not in st.session_state:
        st.session_state.stock_rates = {}
    if "stock_options" not in st.session_state:
        st.session_state.stock_options = []

    uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx", "xlsm"])
    if uploaded is not None:
        st.session_state.file_bytes = uploaded.getvalue()
        st.success(f"Uploaded: {uploaded.name}")

    if st.session_state.file_bytes:
        if st.button("Read workbook / detect sheets"):
            try:
                st.session_state.sheet_names = get_sheet_names(st.session_state.file_bytes)
                st.success("Workbook read successfully.")
            except Exception as e:
                st.error("Could not read workbook.")
                st.code(traceback.format_exc())

    if not st.session_state.sheet_names:
        st.info("Upload a workbook, then click 'Read workbook / detect sheets'.")
        return

    sheets = st.session_state.sheet_names
    col_a, col_b = st.columns(2)
    with col_a:
        working_sheet = st.selectbox("Working Sheet", sheets, index=sheets.index("DL ANZ ALLOCATION") if "DL ANZ ALLOCATION" in sheets else 0)
    with col_b:
        reference_sheet = st.selectbox("Reference Sheet", sheets, index=sheets.index("PRINT DB") if "PRINT DB" in sheets else 0)

    st.subheader("Mapping")
    with st.form("mapping_form"):
        c1, c2, c3, c4 = st.columns(4)
        name_row = c1.number_input("Name row", 1, 1000, 4)
        size_row = c2.number_input("Size row", 1, 1000, 5)
        stock_row = c3.number_input("Stock/material row", 1, 1000, 6)
        total_qty_row = c4.number_input("Original total qty row", 1, 1000, 7)

        c5, c6, c7 = st.columns(3)
        country_col = c5.text_input("Country column", "I")
        start_col = c6.text_input("Start column", "AC")
        end_col = c7.text_input("End column", "IG")

        c8, c9, c10, c11 = st.columns(4)
        ds_output_row = c8.number_input("DS/SS output row", 1, 2000, 168)
        clean_qty_output_row = c9.number_input("Clean qty output row", 1, 2000, 169)
        sqm_output_row = c10.number_input("SQM output row", 1, 2000, 170)
        price_output_row = c11.number_input("Price output row", 1, 2000, 171)

        ignore_text = st.text_input("Countries to ignore", "NZ")
        applied = st.form_submit_button("Apply Mapping / Refresh Stock List")

    if applied:
        try:
            st.session_state.stock_options = get_values_for_stock_options(
                st.session_state.file_bytes,
                working_sheet,
                int(stock_row),
                parse_col(start_col, "AC"),
                parse_col(end_col, "IG"),
            )
            st.success(f"Found {len(st.session_state.stock_options)} stock/material names from {parse_col(start_col,'AC')} to {parse_col(end_col,'IG')}.")
        except Exception:
            st.error("Could not detect stocks/materials.")
            st.code(traceback.format_exc())

    st.subheader("Reference Sheet Mapping")
    r1, r2, r3, r4, r5, r6 = st.columns(6)
    ref_name_col = r1.text_input("Ref name col", "C")
    ref_size_col = r2.text_input("Ref size col", "E")
    ref_ds_col = r3.text_input("Ref DS/SS col", "F")
    ref_stock_col = r4.text_input("Ref stock col", "G")
    ref_start_row = r5.number_input("Ref start row", 1, 5000, 12)
    ref_end_row = r6.number_input("Ref end row", 1, 5000, 141)

    st.subheader("Stock/material rates")
    stock_options = st.session_state.stock_options
    selected_stocks = st.multiselect("Pick stock/material to calculate SQM and rate", stock_options)

    ds_loading_pct = st.number_input("Double-sided loading %", min_value=0.0, max_value=500.0, value=20.0, step=1.0)

    with st.form("rates_form"):
        new_rates = {}
        if selected_stocks:
            st.write("Enter rate per SQM for selected stock/materials:")
            for stock in selected_stocks:
                new_rates[stock] = st.number_input(f"Rate per SQM — {stock}", min_value=0.0, value=float(st.session_state.stock_rates.get(stock, 0.0)), step=0.10, format="%.2f")
        else:
            st.info("Select one or more stock/material names first.")
        update_rates = st.form_submit_button("Refresh / Update Rates")

    if update_rates:
        st.session_state.stock_rates.update(new_rates)
        st.success("Rates updated for this session.")
        st.json(st.session_state.stock_rates)

    st.download_button(
        "Download stock rate memory JSON",
        data=json.dumps(st.session_state.stock_rates, indent=2),
        file_name="stock_rate_memory.json",
        mime="application/json",
    )

    st.subheader("Generate")
    if st.button("Generate Excel Workbook"):
        try:
            with st.spinner("Generating workbook..."):
                output_bytes = build_workbook(
                    file_bytes=st.session_state.file_bytes,
                    working_sheet=working_sheet,
                    reference_sheet=reference_sheet,
                    name_row=int(name_row),
                    size_row=int(size_row),
                    stock_row=int(stock_row),
                    total_qty_row=int(total_qty_row),
                    country_col=parse_col(country_col, "I"),
                    start_col=parse_col(start_col, "AC"),
                    end_col=parse_col(end_col, "IG"),
                    ignore_countries=[x.strip() for x in ignore_text.split(",") if x.strip()],
                    ds_output_row=int(ds_output_row),
                    clean_qty_output_row=int(clean_qty_output_row),
                    sqm_output_row=int(sqm_output_row),
                    price_output_row=int(price_output_row),
                    ref_name_col=parse_col(ref_name_col, "C"),
                    ref_size_col=parse_col(ref_size_col, "E"),
                    ref_ds_col=parse_col(ref_ds_col, "F"),
                    ref_stock_col=parse_col(ref_stock_col, "G"),
                    ref_start_row=int(ref_start_row),
                    ref_end_row=int(ref_end_row),
                    selected_stocks=selected_stocks,
                    stock_rates=st.session_state.stock_rates,
                    ds_loading_pct=float(ds_loading_pct),
                )
            st.success("Workbook generated.")
            st.download_button(
                "Download Excel Workbook",
                data=output_bytes,
                file_name="formula_fusion_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception:
            st.error("Workbook generation failed.")
            st.code(traceback.format_exc())


try:
    main()
except Exception:
    st.error("App startup failed.")
    st.code(traceback.format_exc())

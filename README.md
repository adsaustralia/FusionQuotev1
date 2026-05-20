# Excel Formula Fusion — Stock Rates Build

## What this build fixes

This build adds a central Excel sheet called `STOCK RATES`.

After you download the generated workbook, you can change a stock/material rate once in:

`STOCK RATES` column B

All price formulas in the working sheet and the `Stock SQM Summary` sheet will update automatically in Excel.

## DS loading rule

DS loading is not applied globally.

The generated price formula checks the DS/SS output row for each item column:

```excel
=IFERROR(IF(UPPER(AR$168)="DS",AR$170*VLOOKUP(AR$6,'STOCK RATES'!$A:$B,2,FALSE)*1.2,AR$170*VLOOKUP(AR$6,'STOCK RATES'!$A:$B,2,FALSE)),0)
```

Meaning:

- DS = apply loading, default 20%
- SS = no loading
- blank/unknown = no loading

## Deployment

Upload only these files to GitHub root:

- `app.py`
- `requirements.txt`
- `README.md`

Streamlit main file path:

```text
app.py
```

## Requirements

Only two dependencies are used to reduce Streamlit Cloud install failures:

```txt
streamlit>=1.31,<2
openpyxl>=3.1.2,<4
```

## Workflow

1. Upload workbook.
2. Click **Read workbook / detect sheets**.
3. Confirm working and reference sheets.
4. Apply mapping / refresh stock list.
5. Pick stock/materials.
6. Enter rates.
7. Click **Refresh / Update Rates**.
8. Generate workbook.
9. In Excel, edit rates in `STOCK RATES` if needed.

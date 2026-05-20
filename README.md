# Excel Formula Fusion - Final JSON Backup Build

## What this build does

- Upload Excel workbook
- Select working/reference sheets
- Generate Excel formulas
- Clean qty = original total qty minus ignored country qty
- DS/SS lookup from reference sheet
- Stock/material lookup from reference sheet
- Central `STOCK RATES` sheet in exported Excel
- Price formulas reference `STOCK RATES`, so changing one rate updates all linked formulas in Excel
- DS loading applies only when DS/SS row equals `DS`
- Detects quantity multipliers:
  - `set of 4` multiplies by 4 and highlights red
  - `1 PACK = 100` multiplies by 100 and highlights red
  - uncertain set/pack wording highlights orange and does not multiply
- Saves stock rates to backend JSON:
  - `data/stock_rates.json`
- Every save creates dated backups:
  - `data/rate_backups/stock_rates_YYYYMMDD_HHMMSS.json`
  - `data/rate_backups/stock_rates_saved_YYYYMMDD_HHMMSS.json`

## Important Streamlit Cloud warning

Streamlit Cloud local files may reset after reboot or redeploy.  
The backend JSON works during normal server use, but it is not as reliable as a database.

Use the app's **Download current stock_rates.json** and backup download regularly.

Long-term best options:
1. Google Sheet rate database
2. SQLite/Postgres database
3. PrintIQ API stock price source

## Deploy

Upload only these files/folders to GitHub:

```text
app.py
requirements.txt
README.md
data/
```

Streamlit main file:

```text
app.py
```

## Requirements

```txt
streamlit>=1.31,<2
openpyxl>=3.1.2,<4
```

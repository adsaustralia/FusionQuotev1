# Excel Formula Fusion - Final Stock Consolidated Summary Build

This build is designed for your HOKA / ASICS DBDL workbook structure.

## Key fixes

- Prevents circular references by protecting input rows.
- Does **not** overwrite rows 4, 5, 6 or 7.
- Row 7 remains the original total quantity row.
- Clean qty formula uses row 7 minus ignored-country qty.
- Store qty range is used only to subtract ignored countries, not to rebuild total qty.
- Stock/material row is read as a calculated value for the UI, but preserved in the workbook.
- Price formulas reference a central `STOCK RATES` sheet.
- `STOCK RATES` now includes consolidated **Total SQM** and **Total Price** per stock.
- Change a stock rate once in `STOCK RATES` and Excel updates all linked prices and consolidated totals.
- DS loading applies only when the DS/SS output row says `DS`.
- Confident quantity multipliers such as `set of 4` and `1 PACK = 100` are flagged red and multiplied.
- Doubtful set/pack wording is flagged orange and not multiplied.
- Saves backend rates to `data/stock_rates.json`.
- Every rate save creates dated backups in `data/rate_backups/`.

## Default ASICS/HOKA mapping

| Field | Default |
|---|---:|
| Name row | 4 |
| Size row | 5 |
| Stock row | 6 |
| Original total qty row | 7 |
| Store qty rows | 8:153 |
| Country column | I |
| DS/SS output row | 168 |
| Clean qty output row | 169 |
| SQM output row | 170 |
| Price output row | 171 |


## STOCK RATES consolidated summary

The generated `STOCK RATES` sheet now contains:

| Column | Purpose |
|---|---|
| A | Stock name |
| B | Rate |
| C | Total SQM for that stock |
| D | Total Price for that stock |
| E | Last generated timestamp |

The consolidated formulas use `SUMIF` across the selected Start Column to End Column. Example:

```excel
=SUMIF('DL ANZ ALLOCATION'!$AC$6:$IG$6,A2,'DL ANZ ALLOCATION'!$AC$171:$IG$171)
```

So after download, if you change the rate in `STOCK RATES` column B, Excel updates the item prices and the consolidated total price automatically.

## Critical formula behaviour

Clean qty formula follows this structure:

```excel
=AC$7-SUMPRODUCT((AC$8:AC$153)*(--(UPPER($I$8:$I$153)="NZ")))
```

Price formula follows this structure:

```excel
=IF(UPPER(AC$168)="DS",AC$170*VLOOKUP(AC$6,'STOCK RATES'!$A:$B,2,FALSE)*1.2,AC$170*VLOOKUP(AC$6,'STOCK RATES'!$A:$B,2,FALSE))
```

## Deploy to Streamlit Cloud

Upload these files/folders to GitHub root:

```text
app.py
requirements.txt
README.md
data/
.streamlit/
```

Streamlit main file path:

```text
app.py
```

## Warning about Streamlit Cloud storage

Local JSON files can reset after redeploy or server restart. Use the app backup/download options regularly. For permanent storage, move rates to Google Sheets, a database, or PrintIQ API later.

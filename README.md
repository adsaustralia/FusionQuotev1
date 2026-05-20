# Excel Formula Fusion - Final Circular Reference + Speed Fix

This build keeps the previous UI and JSON stock-rate backup workflow, and adds safeguards against circular references.

## Important fixes

- Prevents duplicate output rows such as SQM row and Price row being the same.
- Uses lighter read-only workbook loading for stock detection.
- Keeps DS loading only when the DS/SS output row contains `DS`.
- Keeps central `STOCK RATES` sheet so prices can be changed once in Excel.
- Saves stock rates to `data/stock_rates.json` and dated backups to `data/rate_backups/`.

## Requirements

```txt
streamlit>=1.31,<2
openpyxl>=3.1.2,<4
```

# Excel Formula Fusion V2.3 Minimal Stable

This is a recovery/stability build. It avoids risky startup logic so the Streamlit app can load first.

## Deploy

Upload these files to the root of your GitHub repository:

- `app.py`
- `requirements.txt`
- `README.md`

Set Streamlit Cloud main file path to:

```text
app.py
```

## Important

This build intentionally avoids custom CSS and optional components. If this version does not load, the issue is likely one of these:

1. GitHub repo does not have `app.py` at the root.
2. Streamlit Cloud is pointing to the wrong main file path.
3. Old broken files are still in the repo.
4. Dependency install failed in Streamlit Cloud logs.

## Workflow

1. Upload Excel workbook.
2. Click **Read workbook / detect sheets**.
3. Apply mapping / refresh stock list.
4. Select stock/materials.
5. Enter rate per sqm.
6. Click **Refresh / Update Rates**.
7. Click **Generate Excel Workbook**.

## Logic included

- Clean qty = original total row quantity minus ignored-country quantity.
- Stock names are scanned only from Start Column to End Column.
- DS rate loading defaults to 20%.
- Multiplier detection:
  - `set of 4` multiplies by 4 and highlights red.
  - `1 PACK = 100` multiplies by 100 and highlights red.
  - unclear set/pack text highlights orange and does not multiply.

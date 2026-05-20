# Excel Formula Fusion — Install Fix Build

This build is intentionally minimal to fix Streamlit Cloud requirement installation errors.

## Dependencies

Only two packages are required:

```txt
streamlit>=1.31,<2
openpyxl>=3.1.2,<4
```

No pandas, no AgGrid, no optional UI components.

## Streamlit Cloud Setup

1. Upload `app.py`, `requirements.txt`, and `README.md` to the root of your GitHub repository.
2. In Streamlit Cloud, set the main file path to:

```txt
app.py
```

3. Reboot the app from **Manage App**.

## Workflow

1. Upload Excel workbook.
2. Click **Read workbook / detect sheets**.
3. Select working and reference sheets.
4. Confirm mapping values.
5. Select stocks/materials.
6. Enter stock rates.
7. Click **Refresh / Update Rates**.
8. Click **Generate Excel Workbook**.
9. Download output workbook.

## Notes

This version avoids heavy processing on upload. Workbook processing only happens after button clicks.

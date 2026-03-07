# Quote Conversion Follow-up Tool

Simple web app to compare quote lines to order activity and identify likely quote-to-order conversions.

## What it does

- Uploads:
  - **Order Log** Excel file (uses columns D, E, G, O, U)
  - **Quote Summary** Excel file (uses columns A, B, C, AJ, AW, BJ)
- Matches quote lines to order lines by:
  - `customer_id`
  - `part_number`
  - order date on/after quote date and within a conversion window (default 90 days)
- Shows:
  - Quote line detail with conversion flag and linked order
  - Rep summary with conversion rate and converted net sales
- Allows exporting both reports to Excel.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Then open http://localhost:8000.

## Notes

- The app icon is served directly from `assets/app.ico` at `/favicon.ico` to avoid duplicating binary files.
- Future enhancement: include quote line net price once available in source data.

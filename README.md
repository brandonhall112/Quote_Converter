# Quote Conversion Follow-up Tool

This repository is set up so you can build a **click-to-launch Windows app** from **GitHub Actions** without running terminal commands locally.

## What the app does

- Uploads:
  - **Order Log** Excel file (uses columns D, E, G, O, U)
  - **Quote Summary** Excel file (uses columns A, B, C, AJ, AW, BJ)
- Matches quote lines to order lines by:
  - `customer_id`
  - `part_number`
  - order date on/after quote date and within a conversion window (default 90 days)
- Outputs:
  - Quote line detail with conversion flag and linked order
  - Rep summary with conversion rate and converted net sales
  - Downloadable Excel reports

## No-terminal workflow (GitHub Actions)

1. In GitHub, go to **Actions**.
2. Click **Build Windows App**.
3. Click **Run workflow**.
4. Wait for the run to finish.
5. Open the finished run and download artifact **QuoteConverter-windows**.
6. Double-click `QuoteConverter.exe`.
7. Your browser opens automatically to the app.

## App icon

- Uses `assets/app.ico` as the application icon and favicon.

## Future enhancement

- Add quote-line net price once it is available in the quote summary source data.

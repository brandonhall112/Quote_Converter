# Quote Conversion Follow-up Tool

This repository builds a click-to-launch Windows app from GitHub Actions and produces a follow-up workbook aligned to your follow-up template.

## What changed for your workflow

- No date range selectors in the UI.
- No conversion window selector in the UI.
- Analysis period is driven strictly by the uploaded files.
- Output is consolidated at the **quote number** level (not line-by-line follow-up output).
- Output workbook is generated from your **Parts Follow Up Template.xlsx** so formulas/summary logic are preserved.

## Inputs

- Order Log Excel (uses columns D, E, G, O, U)
- Quote Summary Excel (uses columns A, B, C, AJ, AW, BJ)
- Parts Follow Up Template Excel
  - If not uploaded in the form, the app will look for:
  - `assets/Parts Follow Up Template.xlsx`

## Output

- Download: `Parts_Follow_Up_Output.xlsx`
- Uses your template workbook as a base.
- Populates follow-up quote rows by rep/owner tab where possible.
- Keeps formula cells and summary tabs from the original template.

## No-terminal workflow (GitHub Actions)

1. In GitHub, open **Actions**.
2. Run **Build Windows App**.
3. Download `QuoteConverter-windows` artifact.
4. Double-click `QuoteConverter.exe`.
5. Browser opens to the app.

## Troubleshooting

If Codex says it cannot update an externally changed PR, create a **new PR** from latest `main`.

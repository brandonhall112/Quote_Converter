# Quote Conversion Follow-up Tool

This app takes your Order Log + Quote Summary and creates a follow-up workbook using your `Parts Follow Up Template.xlsx`.

## What it does

- No date range selectors in the UI.
- No conversion window selector in the UI.
- Analysis period is controlled by the files you upload.
- Follow-up output is consolidated by **quote number**.
- Output workbook keeps your template layout/formulas.

## Inputs

- Order Log Excel (columns D, E, G, O, U)
- Quote Summary Excel (columns A, B, C, AJ, AW, BJ)
- Parts Follow Up Template Excel
  - If not uploaded in the form, app uses `assets/Parts Follow Up Template.xlsx`

## Output

- Download file: `Parts_Follow_Up_Output.xlsx`

---

## Render setup (super simple)

### 1) Push this repo to GitHub
Make sure your latest code is on `main`.

### 2) In Render, create a new Web Service
- Click **New +**
- Click **Web Service**
- Choose this GitHub repo
- Branch: `main`

### 3) Fill in these exact settings
- **Runtime:** Python
- **Build Command:**
  ```bash
  pip install -r requirements.txt
  ```
- **Start Command:**
  ```bash
  gunicorn app:app
  ```

### 4) Click Deploy
Render will build and start your app.

### 5) Open your app URL
Render gives you a URL like:
- `https://your-app-name.onrender.com`

That link is your always-online app (as long as the Render service is running).

---

## Notes

- `gunicorn` is included in `requirements.txt` for Render.
- `app.py` is set to use Render's `PORT` automatically.
- Local run still works with:
  ```bash
  python app.py
  ```

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

### 2) In Render, create a new service from Blueprint (recommended)
- Click **New +**
- Click **Blueprint**
- Choose this GitHub repo
- Render will read `render.yaml` and auto-fill settings, including Python 3.11.9

If you do not use Blueprint, create a normal Web Service and copy the same settings manually.

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

### 6) If Render still shows Python 3.14, force-sync the version
In your Render service:
- Open **Settings**
- Set **Python Version** to `3.11.9` (if visible)
- In **Environment**, add:
  - Key: `PYTHON_VERSION`
  - Value: `3.11.9`
- Save changes
- Click **Manual Deploy** -> **Clear build cache & deploy**

Why: your error log shows Render building with Python 3.14.3, which causes pandas to compile from source and fail.

---

## Notes

- `gunicorn` is included in `requirements.txt` for Render.
- `runtime.txt` pins Python to `3.11.9` for hosts that honor runtime files.
- `render.yaml` also pins Python to `3.11.9` so Blueprint deploys use the right version automatically.
- `app.py` is set to use Render's `PORT` automatically.
- Local run still works with:
  ```bash
  python app.py
  ```

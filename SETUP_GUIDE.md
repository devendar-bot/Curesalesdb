# SPINE Dashboard – Google Sheet → Vercel (with GitHub)

The app reads **your spreadsheet directly** (SPINE ORDER REVIEW 2026, ID `YOUR_GOOGLE_SHEET_ID`) and writes changes **back into it**. No Excel files are involved.

```
Your Google Sheet  <—  Apps Script (Code.gs, runs as you)  <—  React app on Vercel  <—  code on GitHub
```

| Action in the app | What happens in the Sheet |
|---|---|
| Open / ↻ Refresh | Reads all tabs live |
| **Save to Sheet** (Transaction Data) | Writes **Qty to be ordered** + **Status** (cols S, T) for the rows currently shown. Rows not shown are untouched |
| **Import** (other 9 tabs) | Writes your *manually mapped* columns into that tab (Replace or Append); CHECK formulas are filled down |
| Excel / PDF | Download only – nothing written |
| Every write | Logged in a tab called **SPINE_Log** (time, action, rows, and the previous values for Save to Sheet) |

---

## Step 1 – Backend (5 min)

Use the Google account that can **edit** the spreadsheet (the sharing link being "view only" is fine).

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Delete the sample code, paste all of `apps-script/Code.gs` and replace `YOUR_GOOGLE_SHEET_ID` and `YOUR_SECRET_TOKEN` at the top (use a long random token; never commit real values to GitHub) → **Save**.
3. In the function dropdown choose **`check`** → **Run** → *Review permissions → Allow*.
   Open **Execution log**: you should see 10 lines like `DEL "del usage" rows: 8506`. A line ending `PROBLEM:` means a tab name differs – fix the name in `CFG` at the top of Code.gs.
4. **Deploy → New deployment → ⚙ Select type → Web app**
   * Execute as: **Me**
   * Who has access: **Anyone**
   → **Deploy** → copy the **Web app URL** (ends with `/exec`).
5. Test in a browser tab: `<Web app URL>?action=all&token=<VITE_API_TOKEN from .env>` → JSON starting `{"TD":[[`.

> After editing Code.gs later: **Deploy → Manage deployments → ✏ → Version: New version → Deploy**.

## Step 2 – The `.env` file (1 min)

Copy `.env.example` to `.env` (it is git-ignored, so it never goes to GitHub) and fill in the three values:

```
VITE_API_URL=https://script.google.com/macros/s/AKfy.../exec
VITE_API_TOKEN=YOUR_SECRET_TOKEN
VITE_SHEET_ID=YOUR_GOOGLE_SHEET_ID
```

To use a different token: change it in **both** `Code.gs` (`API_TOKEN`) and `.env`.

## Step 3 – Test locally (optional)

```bash
npm install
npm run dev        # http://localhost:5173
```

## Step 4 – GitHub

```bash
git init
git add .
git commit -m "SPINE dashboard"
git branch -M main
git remote add origin https://github.com/<you>/spine-dashboard.git   # create the empty PRIVATE repo first at github.com/new
git push -u origin main
```
`.env` is in `.gitignore`, so the token is not uploaded to GitHub.

## Step 5 – Vercel + the .env file

1. vercel.com → **Add New → Project** → import the GitHub repo (Framework: **Vite**, auto-detected).
2. Before clicking Deploy, open **Environment Variables**:
   * **Easiest:** click the **Key** box and **paste the whole content of your `.env` file** – Vercel splits it into the 3 variables automatically (or use *Import .env*).
   * Leave Environments = Production, Preview, Development.
3. **Deploy**. Your app: `https://<project>.vercel.app`.
4. Changed a variable later? **Settings → Environment Variables → edit → Deployments → ⋯ → Redeploy**.

Every `git push` redeploys automatically.

## Using it

* **Variables** (Transaction Data tab): Consumption = Usage ÷ *data months* × *months required*; Shortfall = MWH stock − Consumption (or all stock + PO).
* **Click a Cat. No.** → child table: monthly consumption by Del / Outstation / UP / Bangalore + every transaction.
* **Save to Sheet**: filter first if you only want some items written; a confirmation shows the item count.
* Tabs `Copy of PENDING ORDER` and `test` in your spreadsheet are ignored.

## Security

* `VITE_*` values are compiled into the site – the token is a light lock, not real security. Share the Vercel URL only with your team and turn on **Vercel → Settings → Deployment Protection**.
* Because the Web app runs as you, anyone with the token can read/write those tabs. Rotate the token if it leaks (see Step 2).
* Mistaken write? **File → Version history** in Google Sheets, or the previous values stored in `SPINE_Log`.

## Troubleshooting

| Message | Fix |
|---|---|
| `Unauthorized` | Token in Vercel ≠ `API_TOKEN` in Code.gs |
| "Backend did not return JSON" | Web app not deployed as **Anyone**, or you used the `/dev` URL – use `/exec` |
| `Tab not found: …` | Tab name differs – edit `CFG` in Code.gs, new version |
| `Headers "Cat. No.", "Qty to be ordered", "Status" not found` | Row 1 of Transaction Data must keep those headers |
| Changes in Code.gs have no effect | Deploy a **New version** |
| First load takes 3–6 s | Normal (≈17,000 rows over Apps Script) |

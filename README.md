# Sri Manjunatha Garage — Billing

The billing page runs on GitHub Pages. Every bill is saved to **one Google Sheet**, which gets one tab per month (`Oct-2026`, `Nov-2026`, …) plus a `Parts` tab. The app creates these tabs automatically.

## One-time setup

### 1. Google Sheet + Apps Script
1. Create a new Google Sheet, e.g. *Manjunatha Garage Bills*.
2. **Extensions → Apps Script**. Delete the sample code and paste everything from `apps-script/Code.gs`. Save.
3. **Project Settings (⚙) → Script properties → Add**:
   - `PIN`: a PIN the owner will type in the billing page, e.g. `4821`.
   - `START_BILL_NO` (optional): the first bill number, e.g. `130`.
4. **Deploy → New deployment → ⚙ Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, allow the permissions, and copy the **Web app URL** (it ends in `/exec`).

### 2. GitHub Pages
1. Paste the Web app URL into `config.js` (`API_URL`).
2. Push this folder to a GitHub repo. In the repo, go to **Settings → Pages → Deploy from branch → main / (root)**.
3. Open the Pages URL (`https://<user>.github.io/<repo>/`), enter the PIN once, and bookmark the page.

> After editing `Code.gs` later, go to **Deploy → Manage deployments → ✏ → Version: New version → Deploy**. The URL stays the same.

## Daily use
- Type the Particulars, Qty and Rate. Amount, Total and amount in words are calculated automatically. Parts used before are suggested with their last rate. **Enter** moves to the next box.
- **Save & Print** saves the bill to the month's tab and prints it in the bill layout.
- **Past Bills** lets you search by bill no., vehicle no., phone or name. Click a bill to reprint or edit it.
- To get an Excel copy, open the Sheet and use **File → Download → Microsoft Excel (.xlsx)**.
- **Phone or tablet:** open the page and choose **Add to Home Screen**. In Chrome that's the ⋮ menu; in Safari it's the Share button. The page then opens like an app. It works on phone, tablet and laptop.

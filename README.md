# Kiser Invoice Tools — Excel Add-in

An Excel task-pane add-in that automates progressive invoicing for **LDP / LCP / LMP** projects
(Ed Ball Landscape Architecture · client: Kiser). It replaces the old in-sheet input forms with a
side panel, so multiple users can work on the same workbook without overwriting each other.

The add-in is hosted on **GitHub Pages** and loaded into Excel through a manifest, so a single code
update reaches every user automatically.

---

## What it does

The panel has two tabs:

### Invoice Form
| Button | What it does |
|---|---|
| **Run Input Form** | Adds an LDP / LCP / Change Order section into `LDP & LCP - Invoice Worksheet` + `LDP & LCP - Vendor Tracking`. (Payment Request is **not** handled here.) |
| **Get PR (load lines)** | Loads existing line items + `PR#TBB` amounts into the form to build a Payment Request. |
| **Run Update pr (Vendor Tracking)** | Creates the Payment Request column in the Invoice Worksheet **and** mirrors PR columns into Vendor Tracking. |
| **Run Invoice Generate** | Refreshes `PR#TBB` with formulas and creates a frozen (static) invoice tab for each finalised PR. |

### PO Form
| Button | What it does |
|---|---|
| **Run Load Est** | Loads the contract dropdowns from Vendor Tracking. |
| **Run Update Po** | Adds / updates POs under a contract. |
| **Run Load Po to Move** | Loads a contract's POs into the table so you can tick the ones to move. |
| **Run Move Po** | Moves the ticked POs from the Source contract to the Target contract. |
| **Run Contract Adjust** | Adds a negative contract adjustment row under the selected PO. |

---

## ⚠️ Workbook rules — DO NOT change these (or the scripts break)

The scripts find things by **exact sheet names, exact header text, and marker rows**. If any of these
change, the buttons stop working correctly.

### Sheet names (do not rename)
- `LDP & LCP - Invoice Worksheet`
- `LDP & LCP - Vendor Tracking`
- `PR#TBB`
- `PO Input Form`
- `Payments` (optional — used by Generate Invoice)
- `Timesheet`

### `LDP & LCP - Vendor Tracking`
- Keep the **bold `LDP` and `LCP` marker rows** in column **A**. LDP-vs-LCP is decided by these markers —
  **not** by the contract name. Removing/renaming them sends totals to the wrong column.
- Keep header row at **row 5** with exact labels: `Contract`, `Contract Total`, `Contract Cost`,
  `Management Hours`, `PO Number`, `PO Amount`, `PO Adjustments`, `PO Total`, `LDP Total`,
  `PR#TBB`, `Total`, `Notes`.
- Keep the rows: `Client total Contract/Cost`, `Sub-Contractor Total Paid`, `LCP Analysis` / `LDP Analysis`.
- Do **not** delete the `PR#TBB` column — new PR columns are inserted **before** it.

### `LDP & LCP - Invoice Worksheet`
- Row 2 must keep: `Total To date`, `Paid to Date`, `Payment Request`, `Completed to Date`,
  `% Completed`, `Balance to Finish`.
- Keep the bold **`LCP`** marker row in column A (separates LDP rows from LCP rows).
- Keep `Grand - TOTALS`, `SUB - TOTALS - LDP`, `SUB - TOTALS - LCP` rows.
- Keep the `CO#TBB` and `PR#TBB` columns — new Change-Order / PR columns are inserted **before** them.
- Do not delete the `Total To date` or `Paid to Date` columns (they are the insert anchors).

### `PR#TBB`
- Keep the `SUB - TOTALS`, `Invoiced to Date`, and `Total Paid To date` rows.
- Descriptions live in column A (row 7+); value columns use INDEX/MATCH formulas — don't break them.

### `Payments` (if used)
- Columns: **A = Date, B = Payment type, C = Amount**, data from **row 2**.

### General
- Don't change header spelling/case — scripts match the **exact** text.
- Each Payment Request always creates a **new** column (older PRs are never overwritten).
- Finalised PR tabs (PR#1, PR#2 …) are **static snapshots** — only `PR#TBB` stays live.

---

## Develop locally

```bash
npm install
npm start          # builds, starts https://localhost:3000, sideloads into Excel
```

- `src/taskpane/taskpane.html` — the panel UI (CSS is inline so it loads reliably in every Excel webview).
- `src/taskpane/taskpane.ts` — all the logic.
- Build a production bundle: `npm run build` → output in `dist/`.

> CSS note: styles are inline and use **literal colours** (no CSS variables) and avoid `inset`/`gap`,
> because older Excel desktop webviews (IE11/Trident) don't support them.

---

## Deployment & updates

Hosting is automatic via **GitHub Actions → GitHub Pages**:

1. Edit code, commit, and `git push` to `main`.
2. The `Deploy add-in to GitHub Pages` workflow builds and publishes `dist/`.
3. Site URL: **https://dilranjankr.github.io/forms/**
4. Every user gets the update the next time they open the panel — **the manifest does not need to be
   re-distributed** (as long as the URL stays the same).

If a change doesn't show up, it's usually the Office webview cache:
- Close Excel completely.
- Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`.
- Reopen Excel.

---

## Install the add-in for users

Use `manifest.prod.xml` (built from `manifest.xml` with the production URL).

- **Microsoft 365 Admin (recommended):** admin.microsoft.com → Settings → Integrated apps →
  *Upload custom apps* → upload the manifest → assign to users.
- **SharePoint App Catalog:** App Catalog site → *Apps for Office* → upload the manifest →
  users get it via Excel → Insert → Add-ins → *My Organization*.
- **Per user (no admin):** Excel → Insert → My Add-ins → *Upload My Add-in* → choose the manifest.

> Sideloading is per-user/per-machine — it does **not** travel inside the workbook file. To reach all
> users you must deploy via Admin Center or the SharePoint App Catalog.

# Install — Kiser Invoice Tools (Excel add-in)

A one-time setup, ~2 minutes per laptop. After this, the add-in panel opens automatically
whenever you open a Kiser invoice workbook.

---

## What you need
1. **Excel for Microsoft 365** (desktop on Windows / Mac, or Excel on the web).
2. The **`manifest.prod.xml`** file (sent along with this guide).
3. Internet connection.

Save the manifest file somewhere easy to find — e.g. `Documents\Kiser Add-in\manifest.prod.xml`.

---

## Excel for Windows / Mac (desktop)

1. Open **Excel**.
2. Click the **Insert** tab.
3. Click **Get Add-ins** (or **My Add-ins** if "Get Add-ins" isn't visible).
4. In the dialog, click **MY ADD-INS** (top of the dialog).
5. Click **Manage My Add-ins** at the top-right → **Upload My Add-in**.
6. Click **Browse…** and select the `manifest.prod.xml` file you saved.
7. Click **Upload**.

That's it. The **Kiser Invoice Tools** panel will appear on the right side of Excel.

> If you don't see the panel right away, go to **Home tab** → **Show Taskpane** /
> **Kiser Invoice Tools** button.

---

## Excel on the web (browser)

1. Open the workbook in Excel for the web.
2. Click **Insert** → **Office Add-ins**.
3. Click **MY ADD-INS** → **Manage My Add-ins** → **Upload My Add-in**.
4. Choose `manifest.prod.xml` → **Upload**.

---

## After installing

- The add-in shows two main tabs in its side panel: **Invoice Form** and **PO Form**, plus a
  **Help** tab with usage rules.
- Open any Kiser invoice workbook (from SharePoint / OneDrive) and use the panel buttons
  to add LDP / LCP sections, Payment Requests, POs, and to generate invoices.

---

## Troubleshooting

**"Upload My Add-in" button doesn't appear**
- You may be using a very old Excel build. Use Excel from Microsoft 365 (subscription) — the
  perpetual versions (Excel 2016/2019) sometimes block upload.

**Panel says "Loading the add-in…" and never finishes**
- Check internet — the add-in loads from https://dilranjankr.github.io/forms/
- Close Excel completely and reopen.

**Panel loads but icons look black or text isn't styled**
- The Office webview cache is stale. Close Excel, then delete the folder:
  `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
  (Windows + R → paste the path → delete contents → reopen Excel.)

**I uninstalled by mistake**
- Just repeat the install steps above.

---

## Updates — nothing to do 🎉

You don't need to reinstall when the add-in is updated.
The code lives on a hosted page; every time you open the panel, you get the latest version
automatically. The manifest only changes if the hosting URL itself moves.

---

## Need help?
Send a screenshot of the issue to your IT / the add-in maintainer.

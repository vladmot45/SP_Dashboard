# SP Trading Operations — Streamlit App

A data entry and analysis tool for your SP-Stocks Excel workbook.

---

## Setup (one time)

1. Make sure Python is installed (3.9 or newer)
2. Open a terminal / command prompt in this folder
3. Install dependencies:

```
pip install -r requirements.txt
```

---

## Run the app

```
streamlit run app.py
```

Your browser will open automatically at http://localhost:8501

---

## First time use

1. In the sidebar, paste the **full path** to your Excel file.
   Example: `C:\Users\YourName\Documents\SP-Stocks_24_03_2026.xlsx`
2. You should see a green "File connected" message.
3. Navigate using the sidebar menu.

---

## What each page does

| Page | Description |
|------|-------------|
| **Dashboard** | Live exchange rates (NBP), stock positions, S26 contract summary |
| **Contracts (SP)** | Filter/view all SP contracts + form to add new ones |
| **Weighbridge (WAGI)** | View recent weighings + log new truck/wagon weights |
| **Incoming Goods** | Log incoming wagons/containers/trucks to Inc sheets |
| **Exchange Rates** | View NBP rates + one-click write to Ex rate_USD sheet |
| **Raw Data Viewer** | Read-only view of any sheet for quick checks |

---

## Important notes

- The app **writes directly** to your Excel file. Keep a backup copy.
- The Excel file must be **closed in Excel** when the app writes to it, otherwise you'll get a permission error.
- Exchange rates come from the **NBP public API** (no key needed) and are cached for 1 hour.
- The `load_sheet` cache refreshes every 5 minutes, or immediately after any write.

---

## Extending the app

Each page is a clearly labelled block in `app.py` (look for the `# PAGE:` comments).
To add a new form or sheet, copy the pattern from an existing page.

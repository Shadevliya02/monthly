# Paid Monthly Dashboard · Outbound & Inbound

Static web dashboard (no build step for the page itself) for invoices paid in 2026. Outbound = Type of Work "Outbound"; Inbound = "Project Delta" (shown as Type Project "Project Delta" in the invoice table).

Open `index.html` in a browser, or deploy the folder as-is (e.g. Vercel, preset "Other").

## Pages
- **Dashboard**: one-screen overview. Pick `2026` for the full year or a month to zoom in.
  Total paid (Outbound vs Inbound), paid per month, invoices not paid yet (returned / in review),
  average days from submit to paid, payment speed buckets, return history, top partners, and the invoices waiting longest.
- **Invoice data**: the full list with filters (type, partner, status, paid month, search), sorting and paging.
  Sub-menus open it pre-filtered: Outbound, Inbound, Returned, In review. Click a row for the invoice timeline.

## Updating the data
- `data/Order_Outbound_Inbound_2026.xlsx`: base list of all orders (sheet `Orders`).
- `data/PAID_*.xlsx`: paid updates (sheet `Orders`). Rows with an existing invoice number update that order; new invoice numbers are added. Orders not in the update file are kept as they are.
- `data/return_dates.json`: return/resubmit dates, keyed by invoice number.

Run `pip install openpyxl` once, then `python3 tools/build_data.py` to rewrite `data.js`. Commit and push; Vercel redeploys automatically.

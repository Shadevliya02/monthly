# Paid Monthly Dashboard · Outbound & Inbound

Static web dashboard (no build step for the page itself) for invoices paid in 2026. Outbound = Type of Work "Outbound"; Inbound = "Project Delta".

Open `index.html` in a browser, or deploy the folder as-is (e.g. Vercel, preset "Other").

## Pages
- **KPI Dashboard**: one-screen overview. Pick `2026` for the full year or a month to zoom in.
  Total paid (Outbound vs Inbound), paid per month, invoices not paid yet (returned / in review),
  average days from submit to paid, payment speed buckets, return history, top partners, and the invoices waiting longest.
- **Invoice data**: the full list with filters (type, partner, status, paid month, search), sorting and paging.
  Sub-menus open it pre-filtered: Outbound, Inbound, Returned, In review. Click a row for the invoice timeline.

## Updating the data
1. Replace `data/Order_Outbound_Inbound_2026.xlsx` with the new export (sheet `Orders`, same columns).
2. Return/resubmit dates live in `data/return_dates.json`, keyed by invoice number.
3. Run `pip install openpyxl` once, then `python3 tools/build_data.py`. This rewrites `data.js`.
4. Commit and push; Vercel redeploys automatically.

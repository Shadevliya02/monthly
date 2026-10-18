# Paid Monthly Dashboard · Outbound & Inbound

Static web dashboard (no build step for the page itself) for invoices paid in 2026. Outbound = Type of Work "Outbound"; Inbound = "Project Delta" (shown as Type Project "Project Delta" in the invoice table).

Open `index.html` in a browser, or deploy the folder as-is (e.g. Vercel, preset "Other").

## Pages
- **Dashboard**: one screen on a 1920×1080 display. Pick `2026` or a month at the top right.
  Top row: total paid, invoices received, average days to get paid, not paid yet.
  Middle: paid / received per month, payment speed, top partners. Bottom: invoices waiting longest and return history.
  Each tile has a tag showing whether it follows the selected period or shows today's situation, and an ⓘ with a plain explanation.
- **Invoices**: All invoices, by type (Outbound, Inbound), by status (Paid, In review, Returned) and Return history.
  Filters, removable filter chips, sorting, rows per page, sticky header. Click a row for the invoice timeline.
- **Guide & definitions**: how to use the site and what each term means.

## Updating the data
- `data/Order_Outbound_Inbound_2026.xlsx`: base list of all orders (sheet `Orders`).
- `data/updates/*.xlsx`: later exports (sheet `Orders`), applied in file-name order, so start each name with the date (`2026-10-19_...`). Rows with an existing invoice number update that order; new invoice numbers are added. Orders not in the update file are kept as they are.
- `data/return_dates.json`: return/resubmit dates, keyed by invoice number.

Run `pip install openpyxl` once, then `python3 tools/build_data.py` to rewrite `data.js`. Commit and push; Vercel redeploys automatically.

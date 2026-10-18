"""Build data.js for the dashboard from the xlsx exports.

Usage:  python3 tools/build_data.py
Needs:  pip install openpyxl

Sources (in data/):
  Order_Outbound_Inbound_2026.xlsx  base list of all orders (sheet "Orders")
  PAID_*.xlsx                       paid updates (sheet "Orders"); applied in name order.
                                    A row whose invoice number already exists updates that order,
                                    otherwise it is added as a new order.
  return_dates.json                 return/resubmit cycles keyed by invoice number
"""
import datetime
import json
import pathlib
import re

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
BASE = DATA / "Order_Outbound_Inbound_2026.xlsx"
OUT = ROOT / "data.js"

# accepted header names per field (exports are not always consistent)
COLS = {
    "no": ["No.", "No"],
    "work": ["Type of Work"],
    "project": ["Type Project"],
    "partner": ["Partner"],
    "contract": ["Contract Number"],
    "desc": ["Description"],
    "inv": ["Invoice Number"],
    "value": ["BAST Value"],
    "status": ["Current Status", "Status Payment"],
    "submit": ["Submit Invoice Date"],
    "paid": ["Paid Date"],
}


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("​", "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def key(inv):
    return re.sub(r"\s+", "", clean(inv)).upper()


def date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    s = clean(v)
    for fmt in ("%d %B %Y", "%d %b %Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError(f"Unrecognised date: {v!r}")


def read(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Orders"]
    header = [clean(c) for c in next(ws.iter_rows(max_row=1, values_only=True))]
    idx = {}
    for field, names in COLS.items():
        found = [header.index(n) for n in names if n in header]
        if not found:
            raise KeyError(f"{path.name}: column for {field!r} not found (looked for {names})")
        idx[field] = found[0]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not any(c is not None for c in r):
            continue
        rows.append({
            "no": r[idx["no"]],
            "work": clean(r[idx["work"]]),
            "project": clean(r[idx["project"]]),
            "partner": clean(r[idx["partner"]]),
            "contract": clean(r[idx["contract"]]),
            "desc": clean(r[idx["desc"]]),
            "inv": clean(r[idx["inv"]]),
            "value": float(r[idx["value"]] or 0),
            "status": clean(r[idx["status"]]).lower(),
            "submit": date(r[idx["submit"]]),
            "paid": date(r[idx["paid"]]),
        })
    return rows


def main():
    orders = read(BASE)
    by_key = {}
    for o in orders:
        by_key.setdefault(key(o["inv"]), []).append(o)
    next_no = max(o["no"] for o in orders) + 1

    sources = [BASE.name]
    stats = []
    for path in sorted(DATA.glob("PAID_*.xlsx")):
        updated = added = 0
        for u in read(path):
            u["status"] = u["status"] or "paid"
            matches = by_key.get(key(u["inv"]))
            if matches:
                for o in matches:
                    o.update({k: v for k, v in u.items() if k != "no"})
                updated += 1
            else:
                u["no"] = next_no
                next_no += 1
                orders.append(u)
                by_key[key(u["inv"])] = [u]
                added += 1
        sources.append(path.name)
        stats.append(f"{path.name}: {updated} updated, {added} added")

    returns = {key(k): v for k, v in json.loads((DATA / "return_dates.json").read_text()).items() if not k.startswith("_")}
    for o in orders:
        o["cycles"] = returns.get(key(o["inv"]), [])

    payload = {"source": " + ".join(sources), "built": datetime.date.today().isoformat(), "orders": orders}
    OUT.write_text("window.PAYMENT_DATA = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n")
    print(f"Wrote {OUT.name}: {len(orders)} orders, {sum(o['status'] == 'paid' for o in orders)} paid, "
          f"{sum(bool(o['cycles']) for o in orders)} with return dates")
    for s in stats:
        print("  " + s)


if __name__ == "__main__":
    main()

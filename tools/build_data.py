"""Build data.js for the dashboard from the Orders sheet of the xlsx export.

Usage:  python3 tools/build_data.py [path/to/Order_Outbound_Inbound_2026.xlsx]
Needs:  pip install openpyxl
"""
import datetime
import json
import pathlib
import re
import sys

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "Order_Outbound_Inbound_2026.xlsx"
RETURNS = ROOT / "data" / "return_dates.json"
OUT = ROOT / "data.js"

STATUS = {"paid": "paid", "review": "review", "returned": "returned"}


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("​", "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


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


def main():
    ws = openpyxl.load_workbook(SRC, data_only=True)["Orders"]
    header = [clean(c) for c in next(ws.iter_rows(max_row=1, values_only=True))]
    col = {name: i for i, name in enumerate(header)}
    returns = {k: v for k, v in json.loads(RETURNS.read_text()).items() if not k.startswith("_")}

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not any(c is not None for c in r):
            continue
        inv = clean(r[col["Invoice Number"]])
        status = clean(r[col["Current Status"]]).lower()
        rows.append({
            "no": r[col["No."]],
            "work": clean(r[col["Type of Work"]]),
            "project": clean(r[col["Type Project"]]),
            "partner": clean(r[col["Partner"]]),
            "contract": clean(r[col["Contract Number"]]),
            "desc": clean(r[col["Description"]]),
            "inv": inv,
            "value": float(r[col["BAST Value"]] or 0),
            "status": STATUS.get(status, status),
            "submit": date(r[col["Submit Invoice Date"]]),
            "paid": date(r[col["Paid Date"]]),
            "cycles": returns.get(inv, []),
        })

    payload = {
        "source": SRC.name,
        "built": datetime.date.today().isoformat(),
        "orders": rows,
    }
    OUT.write_text("window.PAYMENT_DATA = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n")
    print(f"Wrote {OUT} with {len(rows)} orders ({sum(1 for r in rows if r['cycles'])} with return dates)")


if __name__ == "__main__":
    main()

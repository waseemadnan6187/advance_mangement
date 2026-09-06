import os
from datetime import datetime

import openpyxl

from database import get_connection, get_employee_summary

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "advance.xlsm")

_cache = {}
_CACHE_MTIME = None


def _load_excel():
    global _cache, _CACHE_MTIME
    try:
        mtime = os.path.getmtime(EXCEL_FILE)
    except OSError:
        return None
    if _CACHE_MTIME == mtime and _cache:
        return _cache

    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True, keep_vba=False)
    ws = wb["Data"]
    summary = {}
    integrity = {"balance_err": 0, "running_err": 0, "rows": 0}
    g = 0
    for row in ws.iter_rows(min_row=6, values_only=True):
        date, name = row[2], row[3]
        if date is None or name is None or not isinstance(date, datetime):
            continue
        adv = float(row[5]) if row[5] else 0
        ded = float(row[6]) if row[6] else 0
        bal = float(row[7]) if row[7] else 0
        tbal = float(row[8]) if row[8] else 0

        integrity["rows"] += 1
        if abs(bal - (adv - ded)) > 0.01:
            integrity["balance_err"] += 1
        g += adv - ded
        if abs(tbal - g) > 0.01:
            integrity["running_err"] += 1

        key = str(name).strip().lower()
        s = summary.setdefault(key, {"name": str(name).strip(), "adv": 0.0, "ded": 0.0, "cnt": 0})
        s["adv"] += adv
        s["ded"] += ded
        s["cnt"] += 1

    wb.close()
    _cache = {"summary": summary, "integrity": integrity}
    _CACHE_MTIME = mtime
    return _cache


def _db_counts():
    conn = get_connection()
    rows = conn.execute(
        "SELECT emp_name, COUNT(*) as cnt FROM transactions GROUP BY emp_name"
    ).fetchall()
    conn.close()
    return {r["emp_name"].lower(): r["cnt"] for r in rows}


def run_reconcile():
    excel = _load_excel()
    if excel is None:
        return {"error": f"Source file not found: {EXCEL_FILE}"}

    eb = _db_counts()
    db_summary = {}
    for s in get_employee_summary():
        key = s["name"].lower()
        db_summary[key] = {
            "name": s["name"],
            "adv": s["total_advance"],
            "ded": s["total_deduction"],
            "bal": s["balance"],
            "cnt": eb.get(key, 0),
        }

    xl = excel["summary"]
    keys = set(xl) | set(db_summary)

    rows = []
    for key in keys:
        x = xl.get(key)
        d = db_summary.get(key)
        xa = x["adv"] if x else 0
        xd = x["ded"] if x else 0
        xb = xa - xd
        da = d["adv"] if d else 0
        dd = d["ded"] if d else 0
        db_ = d["bal"] if d else 0

        if x and not d:
            status = "only_excel"
        elif d and not x:
            status = "only_db"
        elif abs(xa - da) > 0.01 or abs(xd - dd) > 0.01:
            status = "difference"
        else:
            status = "matched"

        rows.append({
            "excel_name": x["name"] if x else "-",
            "db_name": d["name"] if d else "-",
            "status": status,
            "xl_adv": xa, "db_adv": da, "diff_adv": xa - da,
            "xl_ded": xd, "db_ded": dd, "diff_ded": xd - dd,
            "xl_bal": xb, "db_bal": db_, "diff_bal": xb - db_,
            "xl_cnt": x["cnt"] if x else 0,
            "db_cnt": d["cnt"] if d else 0,
        })

    rows.sort(key=lambda r: (r["excel_name"] if r["excel_name"] != "-" else r["db_name"]).lower())

    matched = sum(1 for r in rows if r["status"] == "matched")
    difference = sum(1 for r in rows if r["status"] == "difference")
    only_excel = sum(1 for r in rows if r["status"] == "only_excel")
    only_db = sum(1 for r in rows if r["status"] == "only_db")

    total_xl_adv = sum(r["xl_adv"] for r in rows)
    total_xl_ded = sum(r["xl_ded"] for r in rows)
    total_xl_bal = sum(r["xl_bal"] for r in rows)
    total_db_adv = sum(r["db_adv"] for r in rows)
    total_db_ded = sum(r["db_ded"] for r in rows)
    total_db_bal = sum(r["db_bal"] for r in rows)

    return {
        "rows": rows,
        "integrity": excel["integrity"],
        "matched": matched,
        "difference": difference,
        "only_excel": only_excel,
        "only_db": only_db,
        "total_xl_adv": total_xl_adv,
        "total_xl_ded": total_xl_ded,
        "total_xl_bal": total_xl_bal,
        "total_db_adv": total_db_adv,
        "total_db_ded": total_db_ded,
        "total_db_bal": total_db_bal,
        "total_employees": len(rows),
    }
import os
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, flash, Response

from database import (
    init_db, get_employee_summary, get_employee_ledger, get_employee_balance,
    get_employee_names, get_employee_info, add_advance, add_deduction,
    add_new_employee, get_recent_transactions, get_transactions_page,
    get_transaction_years, get_transaction_by_id, get_prev_balance_for_trans,
    update_transaction, delete_transaction,
    count_employees, count_transactions,
)
from reconcile import run_reconcile
from import_data import EXCEL_FILE, read_excel_data, import_from_excel

app = Flask(__name__)
app.secret_key = "advance-mgmt-cps"

init_db()

COMPANY_NAME = "CPS Sargodha"
COMPANY_ADDRESS = "New Civil Line, Sargodha"


def number_to_words(n):
    if n == 0:
        return "Zero"
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    if n < 20:
        return ones[n]
    if n < 100:
        return tens[n // 10] + ("" if n % 10 == 0 else " " + ones[n % 10])
    if n < 1000:
        return ones[n // 100] + " Hundred" + ("" if n % 100 == 0 else " " + number_to_words(n % 100))
    if n < 1000000:
        return number_to_words(n // 1000) + " Thousand" + ("" if n % 1000 == 0 else " " + number_to_words(n % 1000))
    if n < 1000000000:
        return number_to_words(n // 1000000) + " Million" + ("" if n % 1000000 == 0 else " " + number_to_words(n % 1000000))
    return number_to_words(n // 1000000000) + " Billion" + ("" if n % 1000000000 == 0 else " " + number_to_words(n % 1000000000))


def amount_in_words(amount):
    n = int(amount)
    dec = round((amount - n) * 100)
    result = number_to_words(n) + " Rupees"
    if dec > 0:
        result += " and " + number_to_words(dec) + " Paisa"
    return result + " Only"


def generate_voucher_html(emp_name, designation, amount, description, trans_date, prev_balance, new_balance, voucher_no=""):
    amt_words = amount_in_words(amount) + "."
    date_formatted = ""
    try:
        dt = datetime.strptime(trans_date, "%Y-%m-%d")
        date_formatted = dt.strftime("%d-%b-%y")
    except Exception:
        date_formatted = trans_date

    company_lines = [line.strip() for line in COMPANY_ADDRESS.split(",") if line.strip()]
    company_html = "<br/>".join(company_lines)

    return f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ 
        font-family: 'Segoe UI', Arial, sans-serif; 
        font-size: 10pt; 
        color: #2b2b2b; 
        background-color: #f4f6f8; 
        padding: 20px; 
    }}
    .voucher-card {{
        max-width: 680px;
        margin: 0 auto;
        background: #ffffff;
        border: 1px solid #e1e6eb;
        border-radius: 6px;
        padding: 30px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    }}
    /* Header Styling */
    .header {{
        text-align: center;
        border-bottom: 2px solid #1a365d;
        padding-bottom: 12px;
        margin-bottom: 16px;
    }}
    .company-name {{
        font-size: 16pt;
        font-weight: 700;
        color: #1a365d;
        text-transform: uppercase;
        letter-spacing: 0.5pt;
    }}
    .company-details {{
        font-size: 8.5pt;
        color: #64748b;
        margin-top: 2px;
    }}
    .voucher-title {{
        text-align: center;
        font-size: 11pt;
        font-weight: 700;
        color: #0f172a;
        background: #f1f5f9;
        padding: 6px;
        border-radius: 4px;
        letter-spacing: 1pt;
        margin-bottom: 16px;
    }}
    /* Info Grid */
    .info-grid {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px 20px;
        margin-bottom: 16px;
        background-color: #fafafa;
        border: 1px solid #f0f0f0;
        padding: 12px;
        border-radius: 4px;
    }}
    .info-item {{ display: flex; flex-direction: column; }}
    .label {{ font-size: 7.5pt; font-weight: 700; text-transform: uppercase; color: #64748b; margin-bottom: 2px; }}
    .value {{ font-size: 9.5pt; font-weight: 600; color: #1e293b; }}
    
    /* Details Section */
    .details-group {{
        margin-bottom: 16px;
    }}
    .detail-row {{
        display: flex;
        justify-content: space-between;
        padding: 8px 0;
        border-bottom: 1px dashed #e2e8f0;
    }}
    /* Highlighted Amount Section */
    .amount-box {{
        background: #f0f9ff;
        border: 1px solid #bae6fd;
        border-radius: 6px;
        padding: 12px;
        text-align: center;
        margin: 16px 0;
    }}
    .amount-title {{ font-size: 8pt; font-weight: 700; color: #0369a1; text-transform: uppercase; }}
    .amount-val {{ font-size: 15pt; font-weight: 800; color: #0284c7; margin: 2px 0; }}
    .amount-words {{ font-size: 8.5pt; font-style: italic; color: #334155; }}

    /* Signatures Layout */
    .signatures-grid {{
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 12px;
        margin-top: 45px;
    }}
    .sig-box {{
        text-align: center;
        border-top: 1.5px solid #cbd5e1;
        padding-top: 6px;
    }}
    .sig-title {{ font-size: 8pt; font-weight: 700; color: #475569; }}

    @media print {{
        body {{ background: #fff; padding: 0; }}
        .voucher-card {{ box-shadow: none; border: 1px solid #ccc; }}
    }}
</style>
</head>
<body>

<div class="voucher-card">
    <!-- Header -->
    <div class="header">
        <div class="company-name">{COMPANY_NAME}</div>
        <div class="company-details">{company_html}</div>
    </div>

    <div class="voucher-title">ADVANCE CASH PAYMENT VOUCHER</div>

    <!-- Metadata -->
    <div class="info-grid">
        <div class="info-item">
            <span class="label">Date</span>
            <span class="value">{date_formatted}</span>
        </div>
        <div class="info-item">
            <span class="label">Employee Name</span>
            <span class="value">{emp_name}</span>
        </div>
        <div class="info-item">
            <span class="label">Previous Balance</span>
            <span class="value">Rs. {prev_balance:,.0f}/-</span>
        </div>
        <div class="info-item">
            <span class="label">Designation</span>
            <span class="value">{designation}</span>
        </div>
    </div>

    <!-- Description -->
    <div class="details-group">
        <div class="detail-row">
            <span class="label">Purpose Of Advance</span>
            <span class="value">{description}</span>
        </div>
    </div>

    <!-- Amount Banner -->
    <div class="amount-box">
        <div class="amount-title">Required Amount</div>
        <div class="amount-val">Rs. {amount:,.0f}/-</div>
        <div class="amount-words"><strong>In Words:</strong> {amt_words}</div>
    </div>

    <!-- Signatures -->
    <div class="signatures-grid">
        <div class="sig-box">
            <div class="sig-title">Prepared By</div>
        </div>
        <div class="sig-box">
            <div class="sig-title">Recommended By</div>
        </div>
        <div class="sig-box">
            <div class="sig-title">Approved By</div>
        </div>
        <div class="sig-box">
            <div class="sig-title">Received By</div>
        </div>
    </div>
</div>

</body>
</html>

"""


def generate_ledger_html(emp_name, designation, code, ledger_data, total_adv, total_ded, balance):
    today = datetime.now().strftime("%d-%B-%Y %I:%M %p")
    rows_html = ""
    running = 0
    for t in ledger_data:
        running += t["advance"] - t["deduction"]
        adv_cell = f"{t['advance']:,.0f}" if t["advance"] else "-"
        ded_cell = f"{t['deduction']:,.0f}" if t["deduction"] else "-"
        bal_color = "#c62828" if running > 0 else ("#2e7d32" if running < 0 else "#000")
        rows_html += f"""
        <tr>
            <td style="text-align:center; padding:4pt 6pt; border-bottom:1pt solid #ccc; font-size:9.5pt;">{t['id']}</td>
            <td style="padding:4pt 6pt; border-bottom:1pt solid #ccc; font-size:9.5pt;">{t['trans_date']}</td>
            <td style="padding:4pt 6pt; border-bottom:1pt solid #ccc; font-size:9.5pt;">{t['description']}</td>
            <td style="text-align:right; padding:4pt 6pt; border-bottom:1pt solid #ccc; color:#1b5e20; font-weight:bold; font-size:9.5pt;">{adv_cell}</td>
            <td style="text-align:right; padding:4pt 6pt; border-bottom:1pt solid #ccc; color:#e65100; font-weight:bold; font-size:9.5pt;">{ded_cell}</td>
            <td style="text-align:right; padding:4pt 6pt; border-bottom:1pt solid #ccc; font-size:9.5pt;">{t['balance']:,.0f}</td>
            <td style="text-align:right; padding:4pt 6pt; border-bottom:1pt solid #ccc; font-weight:bold; color:{bal_color}; font-size:9.5pt;">{running:,.0f}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html>
<head>
<style>
    body {{ font-family: Arial, Helvetica, sans-serif; font-size: 10.5pt; color: #000; margin: 0; }}
    table {{ border-collapse: collapse; }}
    td {{ padding: 4pt 8pt; }}
    .bar {{ background-color: #1a237e; color: #ffffff; text-align: center; }}
    .bar .co {{ font-size: 18pt; font-weight: bold; letter-spacing: 1pt; }}
    .bar .ad {{ font-size: 9pt; color: #cfd8dc; }}
    .title {{ background-color: #e3f2fd; color: #1a237e; text-align: center; font-size: 12pt; font-weight: bold; letter-spacing: 1pt; }}
    .info {{ background-color: #fafafa; border-bottom: 1pt solid #000; font-weight: bold; }}
    .sum {{ font-size: 11pt; font-weight: bold; text-align: center; }}
    th {{ background-color: #1a237e; color: #ffffff; padding: 5pt 6pt; text-align: left; font-size: 9.5pt; }}
    td.data {{ border-bottom: 1pt solid #ccc; }}
    .foot {{ background-color: #eceff1; text-align: center; font-size: 8pt; color: #555; }}
</style>
</head>
<body>
<table style="width: 100%; border: 1.5pt solid #000;" cellpadding="0" cellspacing="0">
    <tr><td class="bar" style="padding: 12pt 8pt 10pt 8pt;">
        <div class="co">{COMPANY_NAME}</div>
        <div class="ad">{COMPANY_ADDRESS}</div>
    </td></tr>
    <tr><td class="title">{emp_name} - Advance Ledger</td></tr>
    <tr><td class="info">
        Employee: {emp_name} &nbsp;&nbsp;&nbsp; Code: {code} &nbsp;&nbsp;&nbsp; Designation: {designation}
    </td></tr>
</table>
<table style="width: 100%; border-collapse: collapse;" cellpadding="0" cellspacing="0">
    <tr>
        <td class="sum" style="width: 33%; background-color: #e8f5e9; color: #1b5e20; border: 1.5pt solid #000;">Total Advance<br/>Rs. {total_adv:,.0f}/-</td>
        <td class="sum" style="width: 34%; background-color: #fff3e0; color: #e65100; border: 1.5pt solid #000;">Total Deduction<br/>Rs. {total_ded:,.0f}/-</td>
        <td class="sum" style="width: 33%; background-color: #ffebee; color: #c62828; border: 1.5pt solid #000;">Balance<br/>Rs. {balance:,.0f}/-</td>
    </tr>
</table>
<table style="width: 100%; border: 1.5pt solid #000; border-collapse: collapse;" cellpadding="0" cellspacing="0">
    <thead>
        <tr>
            <th style="text-align:center; width:5%;">#</th>
            <th style="width:12%;">Date</th>
            <th>Description</th>
            <th style="text-align:right; width:12%;">Advance</th>
            <th style="text-align:right; width:12%;">Deduction</th>
            <th style="text-align:right; width:12%;">Txn Bal</th>
            <th style="text-align:right; width:14%;">Running Total</th>
        </tr>
    </thead>
    <tbody>
        {rows_html}
    </tbody>
</table>
<table style="width: 100%; border-collapse: collapse;" cellpadding="0" cellspacing="0">
    <tr><td class="foot">Printed on {today} | {len(ledger_data)} transactions | Advance Management System</td></tr>
</table>
</body>
</html>"""


@app.context_processor
def inject_company():
    return {"company_name": COMPANY_NAME, "company_address": COMPANY_ADDRESS}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.route("/")
def dashboard():
    summary = get_employee_summary()
    summary_rows = [dict(s) for s in summary]
    total_emp = len(summary_rows)
    total_adv = sum(s["total_advance"] for s in summary_rows)
    total_ded = sum(s["total_deduction"] for s in summary_rows)
    total_bal = sum(s["balance"] for s in summary_rows)
    return render_template(
        "dashboard.html",
        active="dashboard",
        summary=summary_rows,
        total_emp=total_emp,
        total_adv=total_adv,
        total_ded=total_ded,
        total_bal=total_bal,
    )


@app.route("/employees")
def dashboard_json():
    summary = get_employee_summary()
    return [dict(s) for s in summary]


@app.route("/advance")
def advance_entry():
    employees = get_employee_names()

    q = request.args.get("q", "").strip()
    month = request.args.get("month", "").strip()
    year = request.args.get("year", "").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except (TypeError, ValueError):
        page = 1
    try:
        per_page = int(request.args.get("per_page", 50))
    except (TypeError, ValueError):
        per_page = 50

    rows, total = get_transactions_page(page, per_page, q, month, year)
    pages = max(1, -(-total // per_page))
    page = min(page, pages)

    rows, _ = get_transactions_page(page, per_page, q, month, year)
    recent = [dict(t) for t in rows]

    return render_template(
        "advance_entry.html",
        active="advance",
        employees=employees,
        recent=recent,
        recent_count=total,
        total=total,
        page=page,
        per_page=per_page,
        pages=pages,
        q=q,
        month=month,
        year=year,
        years=get_transaction_years(),
    )


@app.route("/employee")
def ledger():
    employees = get_employee_names()
    return render_template(
        "ledger.html",
        active="ledger",
        employees=employees,
        years=get_transaction_years(),
        month=request.args.get("month", "").strip(),
        year=request.args.get("year", "").strip(),
    )


@app.route("/api/employee/<path:name>")
def employee_data(name):
    info = get_employee_info(name)
    ledger = [dict(t) for t in get_employee_ledger(name)]
    if info is None:
        return {"exists": False}, 404

    month = request.args.get("month", "").strip()
    year = request.args.get("year", "").strip()

    running = 0
    shown = []
    for t in ledger:
        running += t["advance"] - t["deduction"]
        t["running_total"] = running
        if year and t["trans_date"][:4] != year:
            continue
        if month:
            try:
                m = int(month)
                if not (1 <= m <= 12):
                    m = None
            except (TypeError, ValueError):
                m = None
            if m and t["trans_date"][5:7] != f"{m:02d}":
                continue
        shown.append(t)

    total_adv = sum(t["advance"] for t in shown)
    total_ded = sum(t["deduction"] for t in shown)
    balance = total_adv - total_ded
    return {
        "exists": True,
        "info": dict(info),
        "ledger": shown,
        "total_adv": total_adv,
        "total_ded": total_ded,
        "balance": balance,
    }


@app.post("/advance/save")
def advance_save():
    emp_name = request.form.get("emp_name", "").strip()
    amount = request.form.get("amount", "").strip()
    description = request.form.get("description", "").strip()
    trans_date = request.form.get("trans_date", "").strip()
    trans_type = request.form.get("trans_type", "")

    error = None
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        amount = 0

    if not emp_name:
        error = "Please select an employee."
    elif amount <= 0:
        error = "Please enter an amount greater than 0."
    elif not description:
        error = "Please enter a description."
    elif not trans_date:
        error = "Please select a date."

    if error:
        flash(error, "error")
        return redirect(url_for("advance_entry"))

    prev_balance = get_employee_balance(emp_name)
    new_balance = prev_balance

    if trans_type == "Advance Payment":
        add_advance(emp_name, amount, description, trans_date)
        new_balance = prev_balance + amount
        trans = get_recent_transactions(1)
        trans_id = trans[0]["id"] if trans else ""
        flash(f"Advance of Rs. {amount:,.0f} saved for {emp_name}.", "success")
        return redirect(url_for("voucher", trans_id=trans_id))
    else:
        add_deduction(emp_name, amount, description, trans_date)
        new_balance = prev_balance - amount
        flash(f"Deduction of Rs. {amount:,.0f} saved for {emp_name}. New Balance: Rs. {new_balance:,.0f}", "success")
        return redirect(url_for("advance_entry"))


@app.route("/voucher/<int:trans_id>")
def voucher(trans_id):
    trans = get_transaction_by_id(trans_id)
    if not trans or trans["advance"] <= 0:
        return "Voucher can only be printed for advance transactions.", 404

    info = get_employee_info(trans["emp_name"])
    designation = info["designation"] if info else ""
    prev_balance, _, _ = get_prev_balance_for_trans(trans_id)
    new_balance = prev_balance + trans["advance"]

    html = generate_voucher_html(
        trans["emp_name"], designation, trans["advance"],
        trans["description"], trans["trans_date"],
        prev_balance, new_balance, f"ADV-{trans['id']}",
    )
    return Response(html, mimetype="text/html")


@app.route("/api/balance/<path:name>")
def employee_balance(name):
    balance = get_employee_balance(name)
    info = get_employee_info(name)
    return {
        "balance": balance,
        "designation": dict(info)["designation"] if info else "",
    }


@app.route("/api/employees")
def employees_json():
    return get_employee_names()


@app.route("/api/transaction/<int:trans_id>")
def transaction_json(trans_id):
    trans = get_transaction_by_id(trans_id)
    if not trans:
        return {"exists": False}, 404
    data = dict(trans)
    data["trans_type"] = "Advance Payment" if trans["advance"] > 0 else "Deduction / Salary Adj."
    data["amount"] = trans["advance"] if trans["advance"] > 0 else trans["deduction"]
    return {"exists": True, "transaction": data}


@app.post("/transaction/<int:trans_id>/edit")
def transaction_edit(trans_id):
    trans = get_transaction_by_id(trans_id)
    if not trans:
        flash("Transaction not found.", "error")
        return redirect(url_for("advance_entry"))

    emp_name = request.form.get("emp_name", "").strip()
    amount = request.form.get("amount", "").strip()
    description = request.form.get("description", "").strip()
    trans_date = request.form.get("trans_date", "").strip()
    trans_type = request.form.get("trans_type", "")

    error = None
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        amount = 0

    if not emp_name:
        error = "Please select an employee."
    elif amount <= 0:
        error = "Please enter an amount greater than 0."
    elif not description:
        error = "Please enter a description."
    elif not trans_date:
        error = "Please select a date."

    if error:
        flash(error, "error")
        return redirect(url_for("advance_entry"))

    is_advance = trans_type == "Advance Payment"
    advance = amount if is_advance else 0
    deduction = 0 if is_advance else amount

    update_transaction(trans_id, emp_name, trans_date, description, advance, deduction)
    flash(f"Transaction #{trans_id} updated.", "success")
    return redirect(request.referrer or url_for("advance_entry"))


@app.post("/transaction/<int:trans_id>/delete")
def transaction_delete(trans_id):
    trans = get_transaction_by_id(trans_id)
    if not trans:
        flash("Transaction not found.", "error")
    else:
        delete_transaction(trans_id)
        flash(f"Transaction #{trans_id} ({trans['emp_name']}) deleted.", "success")
    return redirect(request.referrer or url_for("advance_entry"))


@app.route("/reconcile")
def reconcile():
    data = run_reconcile()
    if "error" in data:
        flash(data["error"], "error")
        data = {
            "rows": [], "integrity": {"rows": 0, "balance_err": 0, "running_err": 0},
            "matched": 0, "difference": 0, "only_excel": 0, "only_db": 0,
            "total_xl_adv": 0, "total_xl_ded": 0, "total_xl_bal": 0,
            "total_db_adv": 0, "total_db_ded": 0, "total_db_bal": 0,
            "total_employees": 0,
        }
    return render_template("reconcile.html", active="reconcile", data=data)


@app.route("/sync")
def sync_page():
    exists = os.path.exists(EXCEL_FILE)
    data = {
        "active": "sync",
        "exists": exists,
        "excel_file": os.path.basename(EXCEL_FILE),
        "mtime": None,
        "preview": None,
        "integrity": None,
        "db_employees": count_employees(),
        "db_transactions": count_transactions(),
    }
    if not exists:
        flash(f"Source file not found: {EXCEL_FILE}", "error")
        return render_template("sync.html", **data)

    data["mtime"] = datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime("%Y-%m-%d %H:%M:%S")
    try:
        preview = read_excel_data(EXCEL_FILE)
        data["preview"] = {
            "employees": len(preview["employees"]),
            "transactions": len(preview["transactions"]),
        }
    except Exception as e:
        flash(f"Could not read Excel file: {e}", "error")

    rec = run_reconcile()
    if "error" not in rec:
        data["integrity"] = rec["integrity"]

    return render_template("sync.html", **data)


@app.post("/sync")
def sync_run():
    if not os.path.exists(EXCEL_FILE):
        flash(f"Source file not found: {EXCEL_FILE}", "error")
        return redirect(url_for("sync_page"))
    try:
        result = import_from_excel(EXCEL_FILE, mode="replace")
        flash(
            f"Sync complete: {result['employees']} employees and "
            f"{result['transactions']} transactions imported from Excel.",
            "success",
        )
    except Exception as e:
        flash(f"Sync failed: {e}", "error")
    return redirect(url_for("sync_page"))


@app.route("/employees/manage")
def manage_employees():
    summary = [dict(s) for s in get_employee_summary()]
    return render_template(
        "employees.html",
        active="employees",
        employees=summary,
    )


@app.post("/employees/add")
def employees_add():
    name = request.form.get("name", "").strip()
    desig = request.form.get("designation", "").strip()
    if not name:
        flash("Please enter an employee name.", "error")
        return redirect(url_for("manage_employees"))
    code = add_new_employee(name, desig)
    flash(f"Employee '{name}' added with code {code}.", "success")
    return redirect(url_for("manage_employees"))


@app.route("/print/ledger/<path:name>")
def ledger_print(name):
    info = get_employee_info(name)
    ledger = [dict(t) for t in get_employee_ledger(name)]
    if not ledger:
        return "No ledger data to print.", 404
    total_adv = sum(t["advance"] for t in ledger)
    total_ded = sum(t["deduction"] for t in ledger)
    balance = total_adv - total_ded
    designation = dict(info)["designation"] if info else ""
    code = dict(info)["code"] if info else ""
    html = generate_ledger_html(name, designation, code, ledger, total_adv, total_ded, balance)
    return Response(html, mimetype="text/html")


if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
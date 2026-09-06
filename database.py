import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "advance.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.executescript("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code INTEGER UNIQUE,
            name TEXT NOT NULL,
            designation TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_name TEXT NOT NULL,
            trans_date DATE NOT NULL,
            description TEXT DEFAULT '',
            advance REAL DEFAULT 0,
            deduction REAL DEFAULT 0,
            balance REAL DEFAULT 0,
            running_total REAL DEFAULT 0
        );

        CREATE INDEX IF NOT EXISTS idx_trans_name ON transactions(emp_name);
        CREATE INDEX IF NOT EXISTS idx_trans_date ON transactions(trans_date);
    """)
    conn.commit()
    conn.close()


def insert_employee(code, name, designation=""):
    conn = get_connection()
    conn.execute(
        "INSERT OR IGNORE INTO employees (code, name, designation) VALUES (?, ?, ?)",
        (code, name.strip(), designation.strip() if designation else ""),
    )
    conn.commit()
    conn.close()


def insert_transaction(emp_name, trans_date, description, advance, deduction, balance, running_total):
    conn = get_connection()
    conn.execute(
        """INSERT INTO transactions (emp_name, trans_date, description, advance, deduction, balance, running_total)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (emp_name.strip(), trans_date, description.strip() if description else "",
         advance or 0, deduction or 0, balance or 0, running_total or 0),
    )
    conn.commit()
    conn.close()


def insert_transactions_batch(rows):
    conn = get_connection()
    conn.executemany(
        """INSERT INTO transactions (emp_name, trans_date, description, advance, deduction, balance, running_total)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        rows,
    )
    conn.commit()
    conn.close()


def insert_employees_batch(rows):
    conn = get_connection()
    conn.executemany(
        "INSERT OR IGNORE INTO employees (code, name, designation) VALUES (?, ?, ?)",
        rows,
    )
    conn.commit()
    conn.close()


def get_all_employees():
    conn = get_connection()
    rows = conn.execute("SELECT * FROM employees ORDER BY code").fetchall()
    conn.close()
    return rows


def get_employee_names():
    conn = get_connection()
    rows = conn.execute("SELECT name FROM employees ORDER BY name").fetchall()
    conn.close()
    return [r["name"] for r in rows]


def get_employee_summary():
    conn = get_connection()
    rows = conn.execute("""
        SELECT e.code, e.name, e.designation,
               COALESCE(SUM(t.advance), 0) as total_advance,
               COALESCE(SUM(t.deduction), 0) as total_deduction,
               COALESCE(SUM(t.advance), 0) - COALESCE(SUM(t.deduction), 0) as balance
        FROM employees e
        LEFT JOIN transactions t ON e.name = t.emp_name
        GROUP BY e.id
        ORDER BY e.code
    """).fetchall()
    conn.close()
    return rows


def get_employee_ledger(emp_name):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM transactions WHERE emp_name = ? ORDER BY trans_date, id",
        (emp_name,),
    ).fetchall()
    conn.close()
    return rows


def get_employee_balance(emp_name):
    conn = get_connection()
    row = conn.execute(
        """SELECT COALESCE(SUM(advance), 0) - COALESCE(SUM(deduction), 0) as balance
           FROM transactions WHERE emp_name = ?""",
        (emp_name,),
    ).fetchone()
    conn.close()
    return row["balance"] if row else 0


def get_employee_info(emp_name):
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM employees WHERE name = ?", (emp_name,)
    ).fetchone()
    conn.close()
    return row


def add_advance(emp_name, amount, description, trans_date=None):
    if trans_date is None:
        trans_date = datetime.now().strftime("%Y-%m-%d")
    balance = get_employee_balance(emp_name)
    running_total = balance + amount
    insert_transaction(emp_name, trans_date, description, amount, 0, amount, running_total)


def add_deduction(emp_name, amount, description, trans_date=None):
    if trans_date is None:
        trans_date = datetime.now().strftime("%Y-%m-%d")
    balance = get_employee_balance(emp_name)
    running_total = balance - amount
    insert_transaction(emp_name, trans_date, description, 0, amount, amount, running_total)


def search_transactions(query):
    conn = get_connection()
    rows = conn.execute(
        """SELECT * FROM transactions
           WHERE emp_name LIKE ? OR description LIKE ?
           ORDER BY trans_date DESC, id DESC LIMIT 500""",
        (f"%{query}%", f"%{query}%"),
    ).fetchall()
    conn.close()
    return rows


def get_transaction_by_id(trans_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM transactions WHERE id = ?", (trans_id,)).fetchone()
    conn.close()
    return row


def get_prev_balance_for_trans(trans_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM transactions WHERE id = ?", (trans_id,)).fetchone()
    if not row:
        conn.close()
        return 0, "", ""
    emp_name = row["emp_name"]
    trans_date = row["trans_date"]
    rows = conn.execute(
        """SELECT advance, deduction FROM transactions
           WHERE emp_name = ? AND (trans_date < ? OR (trans_date = ? AND id < ?))
           ORDER BY trans_date, id""",
        (emp_name, trans_date, trans_date, trans_id),
    ).fetchall()
    conn.close()
    prev = sum(r["advance"] - r["deduction"] for r in rows)
    return prev, emp_name, row["description"]


def get_total_advance_all():
    conn = get_connection()
    row = conn.execute("SELECT COALESCE(SUM(advance), 0) as total FROM transactions").fetchone()
    conn.close()
    return row["total"]


def get_total_deduction_all():
    conn = get_connection()
    row = conn.execute("SELECT COALESCE(SUM(deduction), 0) as total FROM transactions").fetchone()
    conn.close()
    return row["total"]


def add_new_employee(name, designation="", code=None):
    conn = get_connection()
    if code is None:
        row = conn.execute("SELECT COALESCE(MAX(code), 0) + 1 as next_code FROM employees").fetchone()
        code = row["next_code"]
    conn.execute(
        "INSERT OR IGNORE INTO employees (code, name, designation) VALUES (?, ?, ?)",
        (code, name.strip(), designation.strip()),
    )
    conn.commit()
    conn.close()
    return code


def get_recent_transactions(limit=50):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM transactions ORDER BY trans_date DESC, id DESC LIMIT ?",
        (limit,),
    ).fetchall()
    conn.close()
    return rows


def get_all_transactions():
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM transactions ORDER BY trans_date DESC, id DESC"
    ).fetchall()
    conn.close()
    return rows


def get_transactions_page(page=1, per_page=50, search="", month=None, year=None):
    page = max(1, int(page))
    per_page = max(1, int(per_page))
    offset = (page - 1) * per_page

    conditions = []
    params = []
    if search:
        like = f"%{search}%"
        conditions.append("(emp_name LIKE ? OR description LIKE ?)")
        params += [like, like]
    if year:
        conditions.append("substr(trans_date, 1, 4) = ?")
        params.append(str(year))
    if month:
        try:
            m = int(month)
            if 1 <= m <= 12:
                conditions.append("substr(trans_date, 6, 2) = ?")
                params.append(f"{m:02d}")
        except (TypeError, ValueError):
            pass

    where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    conn = get_connection()
    total = conn.execute(
        f"SELECT COUNT(*) as c FROM transactions {where}", params
    ).fetchone()["c"]
    rows = conn.execute(
        f"""SELECT * FROM transactions {where}
            ORDER BY trans_date DESC, id DESC LIMIT ? OFFSET ?""",
        params + [per_page, offset],
    ).fetchall()
    conn.close()
    return rows, total


def get_transaction_years():
    conn = get_connection()
    rows = conn.execute(
        "SELECT DISTINCT substr(trans_date, 1, 4) AS y FROM transactions ORDER BY y DESC"
    ).fetchall()
    conn.close()
    return [r["y"] for r in rows]


def _recompute_running(emp_name):
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM transactions WHERE emp_name = ? ORDER BY trans_date, id",
        (emp_name,),
    ).fetchall()
    running = 0
    for r in rows:
        running += r["advance"] - r["deduction"]
        conn.execute(
            "UPDATE transactions SET balance = ?, running_total = ? WHERE id = ?",
            (r["advance"] - r["deduction"], running, r["id"]),
        )
    conn.commit()
    conn.close()


def update_transaction(trans_id, emp_name, trans_date, description, advance, deduction):
    conn = get_connection()
    old = conn.execute("SELECT * FROM transactions WHERE id = ?", (trans_id,)).fetchone()
    if not old:
        conn.close()
        return False
    old_emp = old["emp_name"]
    conn.execute(
        """UPDATE transactions
           SET emp_name = ?, trans_date = ?, description = ?, advance = ?, deduction = ?
           WHERE id = ?""",
        (emp_name.strip(), trans_date, description.strip() if description else "",
         advance or 0, deduction or 0, trans_id),
    )
    conn.commit()
    conn.close()
    _recompute_running(old_emp)
    if old_emp != emp_name.strip():
        _recompute_running(emp_name.strip())
    return True


def delete_transaction(trans_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM transactions WHERE id = ?", (trans_id,)).fetchone()
    conn.execute("DELETE FROM transactions WHERE id = ?", (trans_id,))
    conn.commit()
    conn.close()
    if row:
        _recompute_running(row["emp_name"])


def count_employees():
    conn = get_connection()
    n = conn.execute("SELECT COUNT(*) AS c FROM employees").fetchone()["c"]
    conn.close()
    return n


def count_transactions():
    conn = get_connection()
    n = conn.execute("SELECT COUNT(*) AS c FROM transactions").fetchone()["c"]
    conn.close()
    return n


def replace_all_data(employees, transactions):
    """Atomically replace all employees and transactions with the given rows."""
    conn = get_connection()
    try:
        conn.execute("BEGIN")
        conn.execute("DELETE FROM transactions")
        conn.execute("DELETE FROM employees")
        conn.executemany(
            "INSERT OR IGNORE INTO employees (code, name, designation) VALUES (?, ?, ?)",
            [(c, n.strip(), d.strip() if d else "") for c, n, d in employees],
        )
        conn.executemany(
            """INSERT INTO transactions (emp_name, trans_date, description, advance, deduction, balance, running_total)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            transactions,
        )
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

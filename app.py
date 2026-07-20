"""
Fully self-contained Flask application — no React, no Node, no npm,
no build step. Just Python + Jinja templates + Tailwind via CDN.

SETUP
    pip install -r requirements.txt --break-system-packages
    python create_starter_workbook.py     # only once, to create company_data.xlsx
    python app.py

Then open http://localhost:5000 on this PC, or http://<this-PC's-LAN-IP>:5000
from any other PC on the same network.

Login accounts (change these below before real use):
    engineer / changeme1   -> Biomedical Engineer dashboard
    sales    / changeme2   -> Sales Rep dashboard
    exec     / changeme3   -> Executive dashboard

STYLING
Every visible page is a plain HTML file in templates/ with Tailwind
utility classes in the class="..." attributes. Edit those directly,
save, refresh the browser — nothing to build or restart except a
plain `python app.py` restart if you change app.py itself (templates
reload automatically).
"""

import threading
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

import pandas as pd
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash

# ---------------------- CONFIG ----------------------

EXCEL_PATH = {
    "Items": "https://docs.google.com/spreadsheets/d/e/2PACX-1vSL0DnLtYg6nghq54svoVFcQEXC_wU41fte-SqdKDFNeVBUegcbQbZmuxM-cgX3LBhs5-VF9lOSMyef/pub?gid=127060097&single=true&output=csv",
    "Sales": "https://docs.google.com/spreadsheets/d/e/2PACX-1vSL0DnLtYg6nghq54svoVFcQEXC_wU41fte-SqdKDFNeVBUegcbQbZmuxM-cgX3LBhs5-VF9lOSMyef/pub?gid=1935654441&single=true&output=csv",
    "Maintenance": "https://docs.google.com/spreadsheets/d/e/2PACX-1vSL0DnLtYg6nghq54svoVFcQEXC_wU41fte-SqdKDFNeVBUegcbQbZmuxM-cgX3LBhs5-VF9lOSMyef/pub?gid=124639110&single=true&output=csv",
    "Complaints": "https://docs.google.com/spreadsheets/d/e/2PACX-1vSL0DnLtYg6nghq54svoVFcQEXC_wU41fte-SqdKDFNeVBUegcbQbZmuxM-cgX3LBhs5-VF9lOSMyef/pub?gid=1263924337&single=true&output=csv"
}
LOW_STOCK_THRESHOLD = 3
MAINTENANCE_DUE_SOON_DAYS = 3

# Change these passwords before real use.
USERS = {
    "sales": {
        "password_hash": generate_password_hash("changeme1"),
        "role": "sales",
        "name": "Sales",
    },
    "engineer": {
        "password_hash": generate_password_hash("changeme2"),
        "role": "engineer",
        "name": "Maintenance",
    },
    "exec": {
        "password_hash": generate_password_hash("changeme3"),
        "role": "exec",
        "name": "Executive",
    },
}

# ---------------------- EXCEL HELPERS ----------------------

_lock = threading.Lock()  # avoid two requests writing the file at once


def read_all_sheets() -> dict[str, pd.DataFrame]:
    with _lock:
        return {name: pd.read_csv(path, dtype=str) for name, path in EXCEL_PATH.items()}


def write_all_sheets(sheets: dict[str, pd.DataFrame]) -> None:
    with _lock:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)


def df_to_records(df: pd.DataFrame) -> list[dict]:
    return df.fillna("").to_dict(orient="records")


def get_items_with_flags(sheets) -> list[dict]:
    records = df_to_records(sheets["Items"])
    for r in records:
        r["stock"] = int(r["stock"])
        r["low_stock"] = r["stock"] < LOW_STOCK_THRESHOLD
    return records

def get_maintenance_with_flags(sheets) -> list[dict]:
    df = sheets["Maintenance"].sort_values("Next_due_date")
    records = df_to_records(df)
    today = datetime.now().date()
    soon_cutoff = today + timedelta(days=MAINTENANCE_DUE_SOON_DAYS)
    for r in records:
        due = pd.to_datetime(r["Next_due_date"]).date()
        r["overdue"] = due < today
        r["due_soon"] = today <= due <= soon_cutoff
    return records


# ---------------------- APP ----------------------

app = Flask(__name__)
app.secret_key = "change-this-to-a-random-string"  # needed for session/flash to work


def login_required(role=None):
    def decorator(view_fn):
        @wraps(view_fn)
        def wrapped(*args, **kwargs):
            if "role" not in session:
                return redirect(url_for("login"))
            if role and session["role"] != role:
                return redirect(url_for(f"{session['role']}_dashboard"))
            return view_fn(*args, **kwargs)

        return wrapped

    return decorator


# ----- Auth -----

@app.route("/", methods=["GET"])
def index():
    if "role" in session:
        return redirect(url_for(f"{session['role']}_dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        user = USERS.get(username)
        if not user or not check_password_hash(user["password_hash"], password):
            return render_template("login.html", error="Invalid username or password")
        session["role"] = user["role"]
        session["name"] = user["name"]
        return redirect(url_for(f"{user['role']}_dashboard"))
    return render_template("login.html", error=None)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ----- Biomedical engineer dashboard -----

@app.route("/engineer")
@login_required(role="engineer")
def engineer_dashboard():
    sheets = read_all_sheets()
    return render_template(
        "engineer.html",
        items=get_items_with_flags(sheets),
        maintenance=get_maintenance_with_flags(sheets),
        complaints=df_to_records(sheets["Complaints"]),
    )


@app.route("/engineer/complaints/<complaint_id>/update", methods=["POST"])
@login_required(role="engineer")
def update_complaint(complaint_id):
    sheets = read_all_sheets()
    df = sheets["Complaints"]
    match = df["id"] == complaint_id
    if match.any():
        df.loc[match, "status"] = request.form.get("status", "")
        df.loc[match, "resolution_notes"] = request.form.get("resolution_notes", "")
        sheets["Complaints"] = df
        write_all_sheets(sheets)
        flash(f"Complaint {complaint_id} updated.")
    return redirect(url_for("engineer_dashboard"))


# ----- Sales rep dashboard -----

@app.route("/sales")
@login_required(role="sales")
def sales_dashboard():
    sheets = read_all_sheets()
    sales_df = sheets["Sales"].sort_values("date", ascending=False)
    return render_template(
        "sales.html",
        items=get_items_with_flags(sheets),
        sales=df_to_records(sales_df),
    )


@app.route("/sales/log", methods=["POST"])
@login_required(role="sales")
def log_sale():
    sheets = read_all_sheets()
    df = sheets["Sales"]
    next_num = 9000 + len(df) + 1
    new_row = {
        "id": f"S-{next_num}",
        "item": request.form.get("item", ""),
        "client": request.form.get("client", ""),
        "amount": request.form.get("amount", "0"),
        "date": request.form.get("date", datetime.now().strftime("%d-%m-%Y")),
    }
    sheets["Sales"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    write_all_sheets(sheets)
    flash("Sale logged.")
    return redirect(url_for("sales_dashboard"))


# ----- Executive dashboard -----

@app.route("/exec")
@login_required(role="exec")
def exec_dashboard():
    sheets = read_all_sheets()

    sales_df = sheets["Sales"]
    total_revenue = sales_df["amount"].astype(float).sum() if len(sales_df) else 0.0

    complaints_df = sheets["Complaints"]
    open_complaints = int((complaints_df["status"] != "Resolved").sum())

    maintenance_df = sheets["Maintenance"]
    cutoff = (datetime.now() + timedelta(days=7)).strftime("%d-%m-%Y")
    maintenance_due_7d = int((maintenance_df["Next_due_date"] <= cutoff).sum())

    items_df = sheets["Items"]
    low_stock_count = int((items_df["stock"].astype(int) < LOW_STOCK_THRESHOLD).sum())

    return render_template(
        "exec.html",
        summary={
            "total_revenue": total_revenue,
            "open_complaints": open_complaints,
            "maintenance_due_7d": maintenance_due_7d,
            "low_stock_count": low_stock_count,
        },
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

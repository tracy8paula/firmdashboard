"""Login accounts (change these below before real use):
    engineer / changeme1   -> Biomedical Engineer dashboard
    sales    / changeme2   -> Sales Rep dashboard
    exec     / changeme3   -> Executive dashboard
"""

import threading
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

import pandas as pd
from flask import (
    Flask,
    jsonify,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash

# ---------------------- CONFIG ----------------------

EXCEL_PATH = Path(__file__).parent / "company_data.xlsx"

LOW_STOCK_THRESHOLD = 3
MAINTENANCE_DUE_SOON_DAYS = 3
RANGE_DAYS = {"week": 7, "month": 30, "6months": 182}

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
        return pd.read_excel(EXCEL_PATH, sheet_name=None, dtype=str)


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

# ---- Analytics API -----  
@app.route("/api/analytics")
@login_required()
def analytics():
    range_key = request.args.get("range", "week")
    days = RANGE_DAYS.get(range_key, 7)

    sheets = read_all_sheets()
    df = sheets["Sales"].copy()
    df["date"] = pd.to_datetime(df["date"])
    df["amount"] = df["amount"].astype(float)

    cutoff = datetime.now() - timedelta(days=days)
    df = df[df["date"] >= cutoff].sort_values("date")

    top = df.groupby("item")["amount"].sum().sort_values(ascending=False).head(5)
    top_products = [{"item": i, "revenue": round(float(v), 2)} for i, v in top.items()]

    if range_key == "week":
        bucket_key = df["date"].dt.strftime("%Y-%m-%d")
        label_fmt = lambda k: datetime.strptime(k, "%Y-%m-%d").strftime("%a %d")
    elif range_key == "month":
        periods = df["date"].dt.to_period("W")
        bucket_key = periods.astype(str)
        label_lookup = {str(p): p.start_time.strftime("Week of %b %d") for p in periods.unique()}
        label_fmt = lambda k: label_lookup[k]
    else:  # 6months
        bucket_key = df["date"].dt.strftime("%Y-%m")
        label_fmt = lambda k: datetime.strptime(k, "%Y-%m").strftime("%b %Y")

    trend_series = df.groupby(bucket_key)["amount"].sum().sort_index()
    trend = [{"label": label_fmt(k), "revenue": round(float(v), 2)} for k, v in trend_series.items()]

    return jsonify({"top_products": top_products, "trend": trend})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

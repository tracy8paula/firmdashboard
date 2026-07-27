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


# authentication


EXCEL_PATH = Path(__file__).parent / "company_data.xlsx"

ENGINEERS={"Alvarez", "Chen", "Kumara", "Othieno", "Kaggwa"}
LOW_STOCK_THRESHOLD = 3
MAINTENANCE_DUE_SOON_DAYS = 3
RANGE_DAYS = {"week": 7, "month": 30, "6months": 182}
DEFAULT_WARRANTY_MONTHS = 12
SERVICE_CYCLE_MONTHS = 6

# the passwords will be changed before real use
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


# excel sheet readers


_lock = threading.Lock() # added to control the requests


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
    df = sheets["Maintenance"].sort_values("next_due_date")
    records = df_to_records(df)
    today = datetime.now().date()
    soon_cutoff = today + timedelta(days=MAINTENANCE_DUE_SOON_DAYS)
    for r in records:
        due = pd.to_datetime(r["next_due_date"]).date()
        r["overdue"] = due < today
        r["due_soon"] = today <= due <= soon_cutoff
        r["days_remaining"] = (due - today).days

        if r.get("date_supplied"):
            supplied = pd.to_datetime(r["date_supplied"])
            warranty_months = int(r.get("warranty_months") or DEFAULT_WARRANTY_MONTHS)
            warranty_end = (supplied + pd.DateOffset(months=warranty_months)).date()
            r["warranty_end_date"] = warranty_end.strftime("%Y-%m-%d")
            r["warranty_valid"] = today <= warranty_end
        else:
            r["warranty_end_date"] = ""
            r["warranty_valid"] = None
    return records


# flask application 


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


# app routes

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


# Biomedical engineer dashboard 


@app.route("/engineer")
@login_required(role="engineer")
def engineer_dashboard():
    sheets = read_all_sheets()
    return render_template(
        "engineer.html",
        items=get_items_with_flags(sheets),
        maintenance=get_maintenance_with_flags(sheets),
        complaints=df_to_records(sheets["Complaints"]),
        engineers=ENGINEERS,
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


@app.route("/engineer/maintenance/<maintenance_id>/reassign", methods=["POST"])
@login_required(role="engineer")
def reassign_engineer(maintenance_id):
    sheets = read_all_sheets()
    df = sheets["Maintenance"]
    match = df["id"] == maintenance_id
    if match.any():
        df.loc[match, "technician"] = request.form.get("technician", "")
        sheets["Maintenance"] = df
        write_all_sheets(sheets)
        flash("Engineer reassigned.")
    return redirect(url_for("engineer_dashboard"))


@app.route("/engineer/maintenance/<maintenance_id>/service", methods=["GET", "POST"])
@login_required(role="engineer")
def service_maintenance(maintenance_id):
    sheets = read_all_sheets()
    m_df = sheets["Maintenance"]
    match = m_df["id"] == maintenance_id
    if not match.any():
        flash("Maintenance record not found.")
        return redirect(url_for("engineer_dashboard"))
    record = m_df[match].iloc[0].to_dict()

    if request.method == "POST":
        issue_reported = request.form.get("issue_reported", "")
        engineer_assigned = request.form.get("engineer_assigned", "")
        part_used = request.form.get("part_used", "")
        part_qty = int(request.form.get("part_qty", "1") or 1)
        date_completed = request.form.get("date_completed", datetime.now().strftime("%Y-%m-%d"))
        notes = request.form.get("notes", "")

        if part_used:
            items_df = sheets["Items"]
            item_match = items_df["name"] == part_used
            if item_match.any():
                current_stock = items_df.loc[item_match, "stock"].astype(int).iloc[0]
                items_df.loc[item_match, "stock"] = str(max(current_stock - part_qty, 0))
                sheets["Items"] = items_df

        sales_df = sheets["Sales"]
        next_num = 9000 + len(sales_df) + 1
        new_sale = {
            "id": f"S-{next_num}", "item": part_used, "client": record.get("client", ""),
            "amount": "0", "date": date_completed, "quantity": str(part_qty),
            "payment_type": "", "amount_paid": "", "balance_due_date": "", "type": "Maintenance",
        }
        sheets["Sales"] = pd.concat([sales_df, pd.DataFrame([new_sale])], ignore_index=True)

        next_due = (pd.to_datetime(date_completed) + pd.DateOffset(months=SERVICE_CYCLE_MONTHS)).date()
        m_df.loc[match, "last_service_date"] = date_completed
        m_df.loc[match, "next_due_date"] = next_due.strftime("%Y-%m-%d")
        m_df.loc[match, "technician"] = engineer_assigned
        sheets["Maintenance"] = m_df

        reports_df = sheets.get("ServiceReports", pd.DataFrame(columns=[
            "id", "maintenance_id", "item", "client", "issue_reported",
            "engineer_assigned", "part_used", "part_qty", "date_completed", "notes",
        ]))
        report_id = f"R-{len(reports_df) + 1001}"
        new_report = {
            "id": report_id, "maintenance_id": maintenance_id, "item": record.get("item", ""),
            "client": record.get("client", ""), "issue_reported": issue_reported,
            "engineer_assigned": engineer_assigned, "part_used": part_used,
            "part_qty": str(part_qty), "date_completed": date_completed, "notes": notes,
        }
        sheets["ServiceReports"] = pd.concat([reports_df, pd.DataFrame([new_report])], ignore_index=True)

        write_all_sheets(sheets)
        flash("Service completed and report generated.")
        return redirect(url_for("view_service_report", report_id=report_id))

    return render_template("service_form.html", record=record, engineers=ENGINEERS, items=get_items_with_flags(sheets))


@app.route("/engineer/service-report/<report_id>")
@login_required(role="engineer")
def view_service_report(report_id):
    sheets = read_all_sheets()
    reports_df = sheets.get("ServiceReports")
    if reports_df is None or not (reports_df["id"] == report_id).any():
        flash("Report not found.")
        return redirect(url_for("engineer_dashboard"))
    report = reports_df[reports_df["id"] == report_id].iloc[0].to_dict()
    return render_template("service_report.html", report=report)


# Sales rep dashboard 


@app.route("/sales")
@login_required(role="sales")
def sales_dashboard():
    sheets = read_all_sheets()
    sales_df = sheets["Sales"]
    if "type" in sales_df.columns:
        sales_df = sales_df[sales_df["type"] != "Maintenance"]
    sales_df = sales_df.sort_values("date", ascending=False)
    return render_template("sales.html", items=get_items_with_flags(sheets), sales=df_to_records(sales_df))


@app.route("/sales/log", methods=["POST"])
@login_required(role="sales")
def log_sale():
    sheets = read_all_sheets()
    df = sheets["Sales"]
    next_num = 9000 + len(df) + 1

    quantity = int(request.form.get("quantity", "1") or 1)
    payment_type = request.form.get("payment_type", "Cash")
    amount = request.form.get("amount_paid", "0")
    amount_paid = request.form.get("amount_paid") or (amount if payment_type == "Cash" else "0")
    balance_due_date = request.form.get("balance_due_date", "")
    item_name = request.form.get("item", "")

    new_row = {
        "id": f"S-{next_num}", "item": item_name, "client": request.form.get("client", ""),
        "amount_paid": amount, "date": request.form.get("date", datetime.now().strftime("%Y-%m-%d")),
        "quantity": str(quantity), "payment_type": payment_type, "amount_paid": amount_paid,
        "balance_due_date": balance_due_date, "type": "Sale",
    }
    sheets["Sales"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    items_df = sheets["Items"]
    item_match = items_df["name"] == item_name
    if item_match.any():
        current_stock = items_df.loc[item_match, "stock"].astype(int).iloc[0]
        items_df.loc[item_match, "stock"] = str(max(current_stock - quantity, 0))
        sheets["Items"] = items_df

    write_all_sheets(sheets)
    flash("Sale logged.")
    return redirect(url_for("sales_dashboard"))


# Executive dashboard

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
    maintenance_due_7d = int((maintenance_df["next_due_date"] <= cutoff).sum())

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


# Analytics API 

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

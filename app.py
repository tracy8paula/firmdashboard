"""
Login accounts (change before real use):
    engineer / changeme1   >>>
    sales    / changeme2
    exec     / changeme3
    hr       / changeme4
"""

import threading
import uuid
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

import pandas as pd
from flask import (
    Flask, flash, jsonify, redirect, render_template,
    request, session, url_for, send_from_directory,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

# ---------------------- CONFIG ----------------------

EXCEL_PATH = Path(__file__).parent / "company_data.xlsx"
UPLOAD_DIR = Path(__file__).parent / "static" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_LOW_STOCK_THRESHOLD = 5
MAINTENANCE_DUE_SOON_DAYS = 14
RANGE_DAYS = {"week": 7, "month": 30, "6months": 182}
DEFAULT_WARRANTY_MONTHS = 12

SERVICE_INTERVALS = {
    "3_weeks": {"label": "3 weeks", "days": 21},
    "1_month": {"label": "1 month", "days": 30},
    "6_months": {"label": "6 months", "days": 182},
}

USERS = {
    "sales": {"password_hash": generate_password_hash("changeme1"), "role": "sales", "name": "Sales"},
    "engineer": {"password_hash": generate_password_hash("changeme2"), "role": "engineer", "name": "Maintenance"},
    "exec": {"password_hash": generate_password_hash("changeme3"), "role": "exec", "name": "Executive"},
    "hr": {"password_hash": generate_password_hash("changeme4"), "role": "hr", "name": "HR"},
}

SHEET_COLUMNS = {
    "Items": ["id", "name", "model_number","serial_number", "location", "status", "stock", "price", "low_stock_threshold", "preferred_supplier", "supplier_contact"],
    "Maintenance": ["id", "item", "serial_number", "client", "location", "date_supplied", "warranty_months", "last_service_date", "next_due_date", "technician"],
    "Complaints": ["id", "item", "model_number", "serial_number", "client", "contact_name", "phone", "date_opened", "issue_description", "status", "resolution_notes", "call_requested"],
    "Sales": ["id", "item", "serial_number","client", "quantity", "list_price", "agreed_price", "salesperson", "payment_type", "payment_route", "payment_stream", "amount_paid", "balance", "balance_due_date", "sale_date", "service_interval", "type"],
    "Installments": ["id", "sale_id", "installment_number", "due_date", "amount_due", "paid_date", "amount_paid", "status"],
    "SpareParts": ["id", "part_name", "model_number", "serial_number", "removed_from_item", "installed_into_item", "date_moved", "logged_by", "notes"],
    "ServiceReports": ["id", "maintenance_id", "item", "client", "issue_reported", "engineer_assigned", "part_used", "part_qty", "date_completed", "notes", "scan_filename"],
    "Staff": ["id", "name", "role", "phone", "email", "date_joined"],
}

# ---------------------- EXCEL HELPERS ----------------------

_lock = threading.Lock()


def read_all_sheets() -> dict[str, pd.DataFrame]:
    with _lock:
        sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, dtype=str)
    for name, cols in SHEET_COLUMNS.items():
        if name not in sheets:
            sheets[name] = pd.DataFrame(columns=cols)
        else:
            for c in cols:
                if c not in sheets[name].columns:
                    sheets[name][c] = ""
    return sheets


def write_all_sheets(sheets: dict[str, pd.DataFrame]) -> None:
    with _lock:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)


def df_to_records(df: pd.DataFrame) -> list[dict]:
    return df.fillna("").to_dict(orient="records")


def new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:8].upper()}"


def parse_date(value):
    return pd.to_datetime(value, format="mixed").date()


# ---------------------- DERIVED DATA ----------------------

def get_items_with_flags(sheets) -> list[dict]:
    records = df_to_records(sheets["Items"])
    for r in records:
        r["stock"] = int(r["stock"] or 0)
        r["price"] = float(r["price"] or 0)
        threshold = int(r["low_stock_threshold"]) if r.get("low_stock_threshold") else DEFAULT_LOW_STOCK_THRESHOLD
        r["low_stock_threshold"] = threshold
        r["low_stock"] = r["stock"] < threshold
    return records


def get_maintenance_with_flags(sheets) -> list[dict]:
    df = sheets["Maintenance"]
    if len(df) == 0:
        return []
    df = df.sort_values("next_due_date")
    records = df_to_records(df)
    today = datetime.now().date()
    soon_cutoff = today + timedelta(days=MAINTENANCE_DUE_SOON_DAYS)
    for r in records:
        if r.get("next_due_date"):
            due = parse_date(r["next_due_date"])
            r["overdue"] = due < today
            r["due_soon"] = today <= due <= soon_cutoff
            r["days_remaining"] = (due - today).days
        else:
            r["overdue"] = r["due_soon"] = False
            r["days_remaining"] = None

        if r.get("date_supplied"):
            supplied = pd.to_datetime(r["date_supplied"], format="mixed")
            warranty_months = int(r.get("warranty_months") or DEFAULT_WARRANTY_MONTHS)
            warranty_end = (supplied + pd.DateOffset(months=warranty_months)).date()
            r["warranty_end_date"] = warranty_end.strftime("%Y-%m-%d")
            r["warranty_valid"] = today <= warranty_end
        else:
            r["warranty_end_date"] = ""
            r["warranty_valid"] = None
    return records


def get_staff(sheets, role=None) -> list[dict]:
    df = sheets["Staff"]
    if role:
        df = df[df["role"] == role]
    return df_to_records(df)


def get_installments_for_sale(sheets, sale_id) -> list[dict]:
    df = sheets["Installments"]
    return df_to_records(df[df["sale_id"] == sale_id].sort_values("installment_number"))


def recompute_sale_balance(sheets, sale_id):
    sales_df = sheets["Sales"]
    match = sales_df["id"] == sale_id
    if not match.any():
        return sheets
    inst_df = sheets["Installments"]
    sale_installments = inst_df[inst_df["sale_id"] == sale_id]
    if len(sale_installments):
        total_paid = sale_installments["amount_paid"].replace("", "0").astype(float).sum()
        sales_df.loc[match, "amount_paid"] = str(total_paid)
    agreed = float(sales_df.loc[match, "agreed_price"].iloc[0] or 0)
    paid = float(sales_df.loc[match, "amount_paid"].iloc[0] or 0)
    sales_df.loc[match, "balance"] = str(agreed - paid)
    sheets["Sales"] = sales_df
    return sheets


# ---------------------- APP ----------------------

app = Flask(__name__)
app.secret_key = "change-this-to-a-random-string"


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
            return render_template("login.html", error="Invalid username or password", usernames=USERS.keys())
        session["role"] = user["role"]
        session["name"] = user["name"]
        return redirect(url_for(f"{user['role']}_dashboard"))
    return render_template("login.html", error=None, usernames=USERS.keys())


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/engineer/inventory")
@login_required(role="engineer")
def inventory_dashboard():
    sheets = read_all_sheets()
    items = get_items_with_flags(sheets)
    low_stock_items = [i for i in items if i["low_stock"]]
    spare_parts = df_to_records(sheets["SpareParts"].sort_values("date_moved", ascending=False)) if len(sheets["SpareParts"]) else []
    return render_template("inventory.html", items=items, low_stock_items=low_stock_items, spare_parts=spare_parts)


@app.route("/engineer/inventory/add", methods=["POST"])
@login_required(role="engineer")
def add_item():
    sheets = read_all_sheets()
    df = sheets["Items"]
    new_row = {
        "id": new_id("I"), "name": request.form.get("name", ""), "location": request.form.get("location", ""),
        "status": request.form.get("status", "Available"), "stock": request.form.get("stock", "0"),
        "price": request.form.get("price", "0"),
        "low_stock_threshold": request.form.get("low_stock_threshold", str(DEFAULT_LOW_STOCK_THRESHOLD)),
        "preferred_supplier": "", "supplier_contact": "",
    }
    sheets["Items"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    write_all_sheets(sheets)
    flash(f"{new_row['name']} added to inventory.")
    return redirect(url_for("inventory_dashboard"))


@app.route("/engineer/inventory/<item_id>/update", methods=["POST"])
@login_required(role="engineer")
def update_item(item_id):
    sheets = read_all_sheets()
    df = sheets["Items"]
    match = df["id"] == item_id
    if match.any():
        df.loc[match, "price"] = request.form.get("price", "0")
        df.loc[match, "low_stock_threshold"] = request.form.get("low_stock_threshold", str(DEFAULT_LOW_STOCK_THRESHOLD))
        df.loc[match, "preferred_supplier"] = request.form.get("preferred_supplier", "")
        df.loc[match, "supplier_contact"] = request.form.get("supplier_contact", "")
        sheets["Items"] = df
        write_all_sheets(sheets)
        flash("Item updated.")
    return redirect(url_for("inventory_dashboard"))


@app.route("/engineer/inventory/spare-parts/log", methods=["POST"])
@login_required(role="engineer")
def log_spare_part():
    sheets = read_all_sheets()
    df = sheets["SpareParts"]
    new_row = {
        "id": new_id("SP"), "part_name": request.form.get("part_name", ""),
        "serial_number": request.form.get("serial_number", ""),
        "removed_from_item": request.form.get("removed_from_item", ""),
        "installed_into_item": request.form.get("installed_into_item", ""),
        "date_moved": request.form.get("date_moved", datetime.now().strftime("%Y-%m-%d")),
        "logged_by": session.get("name", ""), "notes": request.form.get("notes", ""),
    }
    sheets["SpareParts"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    write_all_sheets(sheets)
    flash("Spare part movement logged.")
    return redirect(url_for("inventory_dashboard"))


@app.route("/engineer")
@login_required(role="engineer")
def engineer_dashboard():
    sheets = read_all_sheets()
    return render_template(
        "engineer.html",
        maintenance=get_maintenance_with_flags(sheets),
        complaints=df_to_records(sheets["Complaints"].sort_values("date_opened", ascending=False)) if len(sheets["Complaints"]) else [],
        engineers=get_staff(sheets, role="Engineer"),
    )


@app.route("/engineer/complaints/<complaint_id>/update", methods=["POST"])
@login_required(role="engineer")
def update_complaint(complaint_id):
    sheets = read_all_sheets()
    df = sheets["Complaints"]
    match = df["id"] == complaint_id
    if match.any():
        if df.loc[match, "status"].iloc[0] == "Resolved":
            flash("This complaint is already resolved and locked.")
            return redirect(url_for("engineer_dashboard"))
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
    record = m_df.fillna("")[match].iloc[0].to_dict()

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
                current_stock = int(items_df.loc[item_match, "stock"].iloc[0] or 0)
                items_df.loc[item_match, "stock"] = str(max(current_stock - part_qty, 0))
                sheets["Items"] = items_df

        scan_filename = ""
        uploaded = request.files.get("scan_file")
        if uploaded and uploaded.filename:
            safe_name = secure_filename(uploaded.filename)
            scan_filename = f"{new_id('SCAN')}_{safe_name}"
            uploaded.save(UPLOAD_DIR / scan_filename)

        sales_df = sheets["Sales"]
        new_sale = {
            "id": new_id("S"), "item": part_used, "client": record.get("client", ""),
            "quantity": str(part_qty), "list_price": "0", "agreed_price": "0",
            "salesperson": "", "payment_type": "", "amount_paid": "0", "balance": "0",
            "balance_due_date": "", "sale_date": date_completed, "service_interval": "", "type": "Maintenance",
        }
        sheets["Sales"] = pd.concat([sales_df, pd.DataFrame([new_sale])], ignore_index=True)

        next_due = (pd.to_datetime(date_completed) + pd.DateOffset(months=6)).date()
        m_df.loc[match, "last_service_date"] = date_completed
        m_df.loc[match, "next_due_date"] = next_due.strftime("%Y-%m-%d")
        m_df.loc[match, "technician"] = engineer_assigned
        sheets["Maintenance"] = m_df

        reports_df = sheets["ServiceReports"]
        report_id = new_id("R")
        new_report = {
            "id": report_id, "maintenance_id": maintenance_id, "item": record.get("item", ""),
            "client": record.get("client", ""), "issue_reported": issue_reported,
            "engineer_assigned": engineer_assigned, "part_used": part_used,
            "part_qty": str(part_qty), "date_completed": date_completed, "notes": notes,
            "scan_filename": scan_filename,
        }
        sheets["ServiceReports"] = pd.concat([reports_df, pd.DataFrame([new_report])], ignore_index=True)

        write_all_sheets(sheets)
        flash("Service completed and report generated.")
        return redirect(url_for("view_service_report", report_id=report_id))

    items = get_items_with_flags(sheets)
    engineers = get_staff(sheets, role="Engineer")
    due = parse_date(record["next_due_date"]) if record.get("next_due_date") else None
    today = datetime.now().date()
    if due and (due - today).days > MAINTENANCE_DUE_SOON_DAYS:
        flash("This item isn't due for service yet.")
        return redirect(url_for("engineer_dashboard"))

    return render_template("service_form.html", record=record, engineers=engineers, items=items)


@app.route("/engineer/service-report/<report_id>")
@login_required()
def view_service_report(report_id):
    sheets = read_all_sheets()
    reports_df = sheets["ServiceReports"]
    if not len(reports_df) or not (reports_df["id"] == report_id).any():
        flash("Report not found.")
        return redirect(url_for(f"{session['role']}_dashboard"))
    report = reports_df.fillna("")[reports_df["id"] == report_id].iloc[0].to_dict()
    return render_template("service_report.html", report=report)


@app.route("/uploads/<path:filename>")
@login_required()
def uploaded_file(filename):
    return send_from_directory(UPLOAD_DIR, filename)


@app.route("/complaint", methods=["GET", "POST"])
def public_complaint():
    if request.method == "POST":
        sheets = read_all_sheets()
        df = sheets["Complaints"]
        complaint_id = new_id("C")
        new_row = {
            "id": complaint_id, "item": request.form.get("item", ""),
            "model_number": request.form.get("model_number", ""),
            "serial_number": request.form.get("serial_number", ""),
            "client": request.form.get("client", ""), "contact_name": request.form.get("contact_name", ""),
            "phone": request.form.get("phone", ""), "date_opened": datetime.now().strftime("%Y-%m-%d"),
            "issue_description": request.form.get("issue_description", ""), "status": "Open",
            "resolution_notes": "", "call_requested": "Yes" if request.form.get("call_requested") else "No",
        }
        sheets["Complaints"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        write_all_sheets(sheets)

        contact = None
        m_df = sheets["Maintenance"]
        m_match = m_df[m_df["item"] == new_row["item"]]
        staff_df = sheets["Staff"]
        if len(m_match) and m_match.iloc[0]["technician"]:
            eng_name = m_match.iloc[0]["technician"]
            staff_match = staff_df[staff_df["name"] == eng_name]
            if len(staff_match):
                contact = {"name": eng_name, "phone": staff_match.iloc[0]["phone"], "role": "Engineer"}
        if contact is None:
            s_df = sheets["Sales"]
            s_match = s_df[(s_df["item"] == new_row["item"]) & (s_df["client"] == new_row["client"])]
            if len(s_match) and s_match.iloc[0]["salesperson"]:
                rep_name = s_match.iloc[0]["salesperson"]
                staff_match = staff_df[staff_df["name"] == rep_name]
                if len(staff_match):
                    contact = {"name": rep_name, "phone": staff_match.iloc[0]["phone"], "role": "Sales rep"}

        return render_template("complaint_thanks.html", contact=contact, call_requested=new_row["call_requested"] == "Yes")

    sheets = read_all_sheets()
    items = df_to_records(sheets["Items"])
    return render_template("complaint_form.html", items=items)


@app.route("/sales")
@login_required(role="sales")
def sales_dashboard():
    sheets = read_all_sheets()
    sales_df = sheets["Sales"]
    sales_df = sales_df[sales_df["type"] != "Maintenance"] if len(sales_df) else sales_df
    sales_df = sales_df.sort_values("sale_date", ascending=False) if len(sales_df) else sales_df
    sales = df_to_records(sales_df)
    for s in sales:
        s["installments"] = get_installments_for_sale(sheets, s["id"]) if s.get("payment_type") == "Installments" else []
    return render_template("sales.html", items=get_items_with_flags(sheets), sales=sales, service_intervals=SERVICE_INTERVALS, today=datetime.now().strftime("%Y-%m-%d"))


@app.route("/sales/log", methods=["POST"])
@login_required(role="sales")
def log_sale():
    sheets = read_all_sheets()
    df = sheets["Sales"]

    item_name = request.form.get("item", "")
    quantity = int(request.form.get("quantity", "1") or 1)
    agreed_price = request.form.get("agreed_price", "0")
    payment_type = request.form.get("payment_type", "Cash")
    payment_route = request.form.get("payment_route", "Lumpsum")
    payment_stream = request.form.get("payment_stream", "Bank Transfer")
    client = request.form.get("client", "")
    sale_date = request.form.get("sale_date", datetime.now().strftime("%Y-%m-%d"))
    service_interval = request.form.get("service_interval", "6_months")

    items_df = sheets["Items"]
    item_match = items_df["name"] == item_name
    list_price = float(items_df.loc[item_match, "price"].iloc[0]) if item_match.any() else 0.0

    amount_paid = agreed_price if payment_type == "Cash" else request.form.get("amount_paid", "0")

    sale_id = new_id("S")
    balance = float(agreed_price or 0) - float(amount_paid or 0)

    new_row = {
        "id": sale_id, "item": item_name, "client": client, "quantity": str(quantity),
        "list_price": str(list_price), "agreed_price": agreed_price,
        "salesperson": request.form.get("salesperson", ""), "payment_type": payment_type,
        "amount_paid": amount_paid, "balance": str(balance),
        "balance_due_date": request.form.get("balance_due_date", ""),
        "sale_date": sale_date, "service_interval": service_interval, "type": "Sale",
    }
    sheets["Sales"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    if item_match.any():
        current_stock = int(items_df.loc[item_match, "stock"].iloc[0] or 0)
        items_df.loc[item_match, "stock"] = str(max(current_stock - quantity, 0))
        sheets["Items"] = items_df

    if payment_type == "Installments":
        num_installments = int(request.form.get("num_installments", "2") or 2)
        remaining = float(agreed_price or 0) - float(amount_paid or 0)
        per_installment = round(remaining / num_installments, 2) if num_installments else 0
        inst_df = sheets["Installments"]
        base_date = datetime.strptime(sale_date, "%Y-%m-%d")
        new_installments = []
        for n in range(1, num_installments + 1):
            due = (base_date + pd.DateOffset(months=n)).strftime("%Y-%m-%d")
            new_installments.append({
                "id": new_id("INST"), "sale_id": sale_id, "installment_number": str(n),
                "due_date": due, "amount_due": str(per_installment),
                "paid_date": "", "amount_paid": "0", "status": "Pending",
            })
        sheets["Installments"] = pd.concat([inst_df, pd.DataFrame(new_installments)], ignore_index=True)

    interval_days = SERVICE_INTERVALS.get(service_interval, SERVICE_INTERVALS["6_months"])["days"]
    next_due = (datetime.strptime(sale_date, "%Y-%m-%d") + timedelta(days=interval_days)).strftime("%Y-%m-%d")
    m_df = sheets["Maintenance"]
    new_maintenance = {
        "id": new_id("M"), "item": item_name, "client": client, "location": "Not yet assigned",
        "date_supplied": sale_date, "warranty_months": str(DEFAULT_WARRANTY_MONTHS),
        "last_service_date": "", "next_due_date": next_due, "technician": "",
    }
    sheets["Maintenance"] = pd.concat([m_df, pd.DataFrame([new_maintenance])], ignore_index=True)

    write_all_sheets(sheets)
    flash("Sale logged.")
    return redirect(url_for("sales_dashboard"))


@app.route("/sales/installment/<installment_id>/pay", methods=["POST"])
@login_required(role="sales")
def pay_installment(installment_id):
    sheets = read_all_sheets()
    df = sheets["Installments"]
    match = df["id"] == installment_id
    if match.any():
        df.loc[match, "paid_date"] = request.form.get("paid_date", datetime.now().strftime("%Y-%m-%d"))
        df.loc[match, "amount_paid"] = request.form.get("amount_paid", "0")
        df.loc[match, "status"] = "Paid"
        sheets["Installments"] = df
        sale_id = df.loc[match, "sale_id"].iloc[0]
        sheets = recompute_sale_balance(sheets, sale_id)
        write_all_sheets(sheets)
        flash("Installment payment recorded.")
    return redirect(url_for("sales_dashboard"))


@app.route("/exec")
@login_required(role="exec")
def exec_dashboard():
    sheets = read_all_sheets()

    sales_df = sheets["Sales"].copy()
    client_sales_df = sales_df[sales_df["type"] != "Maintenance"] if len(sales_df) else sales_df
    total_revenue = client_sales_df["agreed_price"].astype(float).sum() if len(client_sales_df) else 0.0

    complaints_df = sheets["Complaints"]
    open_complaints = int((complaints_df["status"] != "Resolved").sum()) if len(complaints_df) else 0

    maintenance = get_maintenance_with_flags(sheets)
    maintenance_due_soon = [m for m in maintenance if m["overdue"] or m["due_soon"]]

    items = get_items_with_flags(sheets)
    low_stock_count = sum(1 for i in items if i["low_stock"])

    status_counts = complaints_df["status"].value_counts().to_dict() if len(complaints_df) else {}
    complaints_by_status = {
        "Open": int(status_counts.get("Open", 0)),
        "In Progress": int(status_counts.get("In Progress", 0)),
        "Resolved": int(status_counts.get("Resolved", 0)),
    }

    revenue_by_item = (
        client_sales_df.groupby("item")["agreed_price"].apply(lambda s: s.astype(float).sum()).sort_values(ascending=False)
        if len(client_sales_df) else pd.Series(dtype=float)
    )
    revenue_streams = [{"item": i, "revenue": round(float(v), 2)} for i, v in revenue_by_item.items()]

    activity = []
    for _, row in client_sales_df.iterrows():
        activity.append({"date": row["sale_date"], "type": "Sale", "description": f"{row['item']} sold to {row['client']}", "status": "Completed"})
    reports_df = sheets["ServiceReports"]
    for _, row in reports_df.iterrows():
        activity.append({"date": row["date_completed"], "type": "Maintenance", "description": f"Serviced {row['item']} for {row['client']}", "status": "Completed"})
    for _, row in complaints_df.iterrows():
        activity.append({"date": row["date_opened"], "type": "Complaint", "description": f"{row['item']} — {row['client']}", "status": row["status"]})
    activity.sort(key=lambda a: a["date"], reverse=True)

    service_reports = df_to_records(reports_df.sort_values("date_completed", ascending=False)) if len(reports_df) else []

    return render_template(
        "exec.html",
        summary={
            "total_revenue": total_revenue, "open_complaints": open_complaints,
            "maintenance_due_7d": len(maintenance_due_soon), "low_stock_count": low_stock_count,
        },
        complaints_by_status=complaints_by_status, recent_activity=activity[:8],
        service_reports=service_reports, revenue_streams=revenue_streams,
        maintenance_due_soon=maintenance_due_soon, exec_name=session.get("name", "Executive"),
    )


@app.route("/api/analytics")
@login_required()
def analytics():
    range_key = request.args.get("range", "week")
    days = RANGE_DAYS.get(range_key, 7)

    sheets = read_all_sheets()
    df = sheets["Sales"].copy()
    df = df[df["type"] != "Maintenance"] if len(df) else df
    if not len(df):
        return jsonify({"top_products": [], "trend": []})

    df["date"] = pd.to_datetime(df["date"], format="mixed")
    df["agreed_price"] = df["agreed_price"].astype(float)

    cutoff = datetime.now() - timedelta(days=days)
    df = df[df["date"] >= cutoff].sort_values("date")

    top = df.groupby("item")["agreed_price"].sum().sort_values(ascending=False).head(5)
    top_products = [{"item": i, "revenue": round(float(v), 2)} for i, v in top.items()]

    if not len(df):
        return jsonify({"top_products": top_products, "trend": []})

    if range_key == "week":
        bucket_key = df["date"].dt.strftime("%Y-%m-%d")
        label_fmt = lambda k: datetime.strptime(k, "%Y-%m-%d").strftime("%a %d")
    elif range_key == "month":
        periods = df["date"].dt.to_period("W")
        bucket_key = periods.astype(str)
        label_lookup = {str(p): p.start_time.strftime("Week of %b %d") for p in periods.unique()}
        label_fmt = lambda k: label_lookup[k]
    else:
        bucket_key = df["date"].dt.strftime("%Y-%m")
        label_fmt = lambda k: datetime.strptime(k, "%Y-%m").strftime("%b %Y")

    trend_series = df.groupby(bucket_key)["agreed_price"].sum().sort_index()
    trend = [{"label": label_fmt(k), "revenue": round(float(v), 2)} for k, v in trend_series.items()]

    return jsonify({"top_products": top_products, "trend": trend})


@app.route("/hr")
@login_required(role="hr")
def hr_dashboard():
    sheets = read_all_sheets()
    staff = get_staff(sheets)

    sales_df = sheets["Sales"]
    client_sales_df = sales_df[sales_df["type"] != "Maintenance"] if len(sales_df) else sales_df
    sales_performance = (
        client_sales_df.groupby("salesperson")["agreed_price"].apply(lambda s: s.astype(float).sum()).sort_values(ascending=False)
        if len(client_sales_df) else pd.Series(dtype=float)
    )
    sales_performance = [{"name": n, "revenue": round(float(v), 2)} for n, v in sales_performance.items() if n]

    reports_df = sheets["ServiceReports"]
    service_performance = (
        reports_df.groupby("engineer_assigned").size().sort_values(ascending=False)
        if len(reports_df) else pd.Series(dtype=int)
    )
    service_performance = [{"name": n, "services_completed": int(v)} for n, v in service_performance.items() if n]

    return render_template("hr.html", staff=staff, sales_performance=sales_performance, service_performance=service_performance)


@app.route("/hr/staff/add", methods=["POST"])
@login_required(role="hr")
def add_staff():
    sheets = read_all_sheets()
    df = sheets["Staff"]
    new_row = {
        "id": new_id("ST"), "name": request.form.get("name", ""), "role": request.form.get("role", ""),
        "phone": request.form.get("phone", ""), "email": request.form.get("email", ""),
        "date_joined": request.form.get("date_joined", datetime.now().strftime("%Y-%m-%d")),
    }
    sheets["Staff"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    write_all_sheets(sheets)
    flash(f"{new_row['name']} added to staff.")
    return redirect(url_for("hr_dashboard"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

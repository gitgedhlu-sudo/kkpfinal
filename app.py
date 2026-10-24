from flask import Flask, render_template, request, jsonify, send_file, abort, session, redirect, url_for
from datetime import datetime
from pathlib import Path
import json
import os
import sqlite3
from functools import wraps
import secrets
import re
import shutil
from urllib.parse import quote_plus
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = Path(__file__).resolve().parent
PRODUCTS_FILE = BASE_DIR / "products.json"
# On Railway, attach a Volume at /app/data. For local development, data is stored
# in ./data unless ORDER_DATA_DIR is explicitly set.
DATA_DIR = Path(os.environ.get("ORDER_DATA_DIR", str(BASE_DIR / "data"))).resolve()
ORDERS_DIR = DATA_DIR / "orders"
ALL_ORDERS_DIR = DATA_DIR / "all_orders"
DB_FILE = DATA_DIR / "orders.sqlite3"
DATA_DIR.mkdir(parents=True, exist_ok=True)
ORDERS_DIR.mkdir(parents=True, exist_ok=True)
ALL_ORDERS_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder="templates", static_folder="public", static_url_path="")
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or secrets.token_hex(32)
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax")
if os.environ.get("RAILWAY_PUBLIC_DOMAIN") or os.environ.get("RAILWAY_ENVIRONMENT"):
    app.config["SESSION_COOKIE_SECURE"] = True

# Initial credentials requested by the site owner. Override via Railway Variables
# before public launch; passwords are configurable without editing source code.
LOGIN_USERS = {
    os.environ.get("LOGIN_ADMIN_USER", "admin"): {"password": os.environ.get("LOGIN_ADMIN_PASSWORD", "dhavalpatel"), "role": "admin"},
    os.environ.get("LOGIN_USER_1", "navsariyuvakmandalpurv"): {"password": os.environ.get("LOGIN_PASSWORD_1", "Navsari@Kkp26!"), "role": "customer"},
    os.environ.get("LOGIN_USER_2", "suratyuvakmandal"): {"password": os.environ.get("LOGIN_PASSWORD_2", "Surat#Kkp84!"), "role": "customer"},
    os.environ.get("LOGIN_USER_3", "valsadyuvukmandal"): {"password": os.environ.get("LOGIN_PASSWORD_3", "Valsad@Kkp73!"), "role": "customer"},
    os.environ.get("LOGIN_USER_4", "vesmayuvukmandal"): {"password": os.environ.get("LOGIN_PASSWORD_4", "vesma@Kkp26!"), "role": "customer"},
}

def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("username"):
            return redirect(url_for("login", next=request.path))
        if session.get("role") != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped


def style_workbook(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="172554")
        cell.alignment = Alignment(vertical="center")
    for column in ws.columns:
        letter = column[0].column_letter
        max_len = max((len(str(c.value or "")) for c in column), default=10)
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 42)


def build_order_excel(order, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Order Receipt"
    ws.append(["KKP FIREWORKS — ORDER RECEIPT"])
    ws.append(["Order number", order["order_id"]])
    ws.append(["Date", order["created_at"].replace("T", " ")])
    ws.append(["Login account", order.get("username", "unknown")])
    ws.append(["Mandal Name", order.get("mandal_name", order["customer_name"])])
    ws.append(["Firm Name", order.get("firm_name", "Not provided")])
    ws.append(["Mobile No", order.get("mobile_no", order.get("phone") or "Not provided")])
    ws.append([])
    ws.append(["Product code", "Product name", "Quantity", "Unit price (INR)", "Amount (INR)"])
    for item in order["items"]:
        ws.append([item["code"], item["name"], int(item["quantity"]), int(item["unit_price"]), int(item["line_total"])])
    ws.append(["", "", "", "TOTAL", int(order["total"])])
    ws.append([])
    ws.append(["Order status", order.get("status", "Order request received")])
    ws.append(["Created by", "Aditya Maganbhai Vasani"])
    ws.merge_cells("A1:E1")
    ws["A1"].font = Font(bold=True, size=15, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="172554")
    ws["A1"].alignment = Alignment(horizontal="center")
    for c in ws[9]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="172554")
    for c in ws[ws.max_row-3]:
        pass
    for row in ws.iter_rows(min_row=10, max_row=ws.max_row):
        for c in row:
            if c.column in (4,5) and isinstance(c.value, (int,float)):
                c.number_format = '₹#,##0'
    for col, width in {"A":18,"B":38,"C":14,"D":20,"E":20}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A10"
    wb.save(path)


def rebuild_all_orders_workbook():
    """Create a consolidated workbook containing every order and its line items."""
    with db_connection() as db:
        rows = db.execute("SELECT order_id, created_at, username, customer_name, phone, total, status, order_details FROM orders ORDER BY CAST(order_id AS INTEGER)").fetchall()
    wb = Workbook()
    summary = wb.active
    summary.title = "All Orders"
    summary.append(["Order number", "Date", "Account", "Customer", "Phone", "Total (INR)", "Status", "Product count"])
    details = wb.create_sheet("Order Items")
    details.append(["Order number", "Date", "Account", "Customer", "Product code", "Product", "Quantity", "Unit price (INR)", "Amount (INR)", "Status"])
    for row in rows:
        order = dict(row)
        data = json.loads(order["order_details"])
        summary.append([order["order_id"], order["created_at"].replace("T", " "), order["username"], order["customer_name"], order["phone"] or "", order["total"], order["status"], len(data.get("items", []))])
        for item in data.get("items", []):
            details.append([order["order_id"], order["created_at"].replace("T", " "), order["username"], order["customer_name"], item["code"], item["name"], item["quantity"], item["unit_price"], item["line_total"], order["status"]])
    style_workbook(summary); style_workbook(details)
    path = ALL_ORDERS_DIR / "all_orders.xlsx"
    wb.save(path)
    return path


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("username"):
            if request.path.startswith("/api/"):
                return jsonify({"success": False, "message": "Please log in to continue."}), 401
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped

with PRODUCTS_FILE.open("r", encoding="utf-8") as f:
    PRODUCTS = json.load(f)


def db_connection():
    connection = sqlite3.connect(DB_FILE, timeout=30)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_database():
    with db_connection() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                order_id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                customer_name TEXT NOT NULL,
                phone TEXT,
                total INTEGER NOT NULL,
                status TEXT NOT NULL,
                pdf_filename TEXT NOT NULL,
                order_details TEXT NOT NULL,
                username TEXT NOT NULL DEFAULT 'unknown'
            )
        """)
        # Migrate existing databases without losing previous orders.
        columns = {row[1] for row in db.execute("PRAGMA table_info(orders)").fetchall()}
        if "username" not in columns:
            db.execute("ALTER TABLE orders ADD COLUMN username TEXT NOT NULL DEFAULT 'unknown'")
        if "phone" not in columns:
            db.execute("ALTER TABLE orders ADD COLUMN phone TEXT")
        # Persistent counter ensures order references are 1, 2, 3... across restarts.
        db.execute("""
            CREATE TABLE IF NOT EXISTS order_sequence (
                sequence_name TEXT PRIMARY KEY,
                last_number INTEGER NOT NULL
            )
        """)
        db.execute("""
            INSERT OR IGNORE INTO order_sequence (sequence_name, last_number)
            VALUES ('orders', 0)
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS inventory (
                product_id TEXT PRIMARY KEY,
                stock INTEGER NOT NULL DEFAULT 100,
                updated_at TEXT NOT NULL
            )
        """)
        # New products start at 100 units. Existing inventory is never overwritten on restart.
        now = datetime.now().astimezone().isoformat(timespec="seconds")
        for product in PRODUCTS:
            db.execute("INSERT OR IGNORE INTO inventory (product_id, stock, updated_at) VALUES (?, 100, ?)",
                       (str(product["id"]), now))


def reserve_order_number():
    # BEGIN IMMEDIATE serializes simultaneous order submissions to avoid duplicates.
    with db_connection() as db:
        db.execute("BEGIN IMMEDIATE")
        row = db.execute(
            "SELECT last_number FROM order_sequence WHERE sequence_name = 'orders'"
        ).fetchone()
        last_number = int(row["last_number"]) if row else 0
        number = last_number + 1
        db.execute(
            "INSERT INTO order_sequence (sequence_name, last_number) VALUES ('orders', ?) "
            "ON CONFLICT(sequence_name) DO UPDATE SET last_number = excluded.last_number",
            (number,)
        )
        db.commit()
        return number

initialize_database()


def escape_paragraph(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_order_pdf(order):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=42, leftMargin=42, topMargin=42, bottomMargin=42)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ShopTitle", parent=styles["Title"], alignment=TA_CENTER,
                              textColor=colors.HexColor("#172554"), spaceAfter=8))
    styles.add(ParagraphStyle(name="Meta", parent=styles["Normal"], leading=17))
    story = [
        Paragraph("KKP FIREWORKS — ORDER RECEIPT", styles["ShopTitle"]),
        Paragraph("Order reference: " + escape_paragraph(order["order_id"]), styles["Meta"]),
        Paragraph("Date: " + escape_paragraph(order["created_at"].replace("T", " ")), styles["Meta"]),
        Paragraph("Mandal Name: " + escape_paragraph(order.get("mandal_name", order["customer_name"])), styles["Meta"]),
        Paragraph("Firm Name: " + escape_paragraph(order.get("firm_name", "Not provided")), styles["Meta"]),
        Paragraph("Mobile No: " + escape_paragraph(order.get("mobile_no", order.get("phone") or "Not provided")), styles["Meta"]),
        Spacer(1, 16),
    ]
    rows = [["Code", "Product", "Qty", "Unit price", "Amount"]]
    for item in order["items"]:
        rows.append([
            str(item["code"]), Paragraph(escape_paragraph(item["name"]), styles["BodyText"]),
            str(item["quantity"]), f'Rs. {item["unit_price"]:,}', f'Rs. {item["line_total"]:,}'
        ])
    rows.append(["", "", "", "TOTAL", f'Rs. {order["total"]:,}'])
    table = Table(rows, colWidths=[42, 220, 42, 75, 75], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#172554")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d8deea")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#f4f7fb")]),
        ("FONTNAME", (3, -1), (-1, -1), "Helvetica-Bold"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.extend([
        table, Spacer(1, 16),
        Paragraph("Order request only — confirm stock and final price with the customer.", styles["Italic"]),
        Spacer(1, 8), Paragraph("Created by Aditya Maganbhai Vasani", styles["Normal"])
    ])
    doc.build(story)
    return buffer.getvalue()


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("username"):
        return redirect(url_for("home"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = LOGIN_USERS.get(username)
        if user and secrets.compare_digest(password, user["password"]):
            session.clear()
            session["username"] = username
            session["role"] = user["role"]
            destination = request.args.get("next", "")
            # Only allow local relative redirects.
            if destination.startswith("/") and not destination.startswith("//"):
                return redirect(destination)
            return redirect(url_for("home"))
        error = "Invalid credentials. Please check your username and password."
    return render_template("login.html", error=error)

@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.get("/")
@login_required
def home():
    categories = list(dict.fromkeys(p["category"] for p in PRODUCTS))
    with db_connection() as db:
        stock_rows = db.execute("SELECT product_id, stock FROM inventory").fetchall()
    stock = {row["product_id"]: row["stock"] for row in stock_rows}
    products_with_stock = [dict(p, stock=stock.get(str(p["id"]), 100)) for p in PRODUCTS]
    return render_template("index.html", products=products_with_stock, categories=categories,
                           is_admin=session.get("role") == "admin")


@app.get("/api/products")
@login_required
def products():
    with db_connection() as db:
        stock_rows = db.execute("SELECT product_id, stock FROM inventory").fetchall()
    stock = {row["product_id"]: row["stock"] for row in stock_rows}
    items = [dict(p, stock=stock.get(str(p["id"]), 100)) for p in PRODUCTS]
    return jsonify({"success": True, "count": len(items), "products": items})


@app.post("/api/orders")
@login_required
def create_order():
    data = request.get_json(silent=True) or {}
    # Accept both the current Mandal/Firm/Mobile form fields and legacy clients.
    mandal_name = str(data.get("mandal_name", data.get("customer_name", ""))).strip()
    firm_name = str(data.get("firm_name", "")).strip()
    phone = str(data.get("mobile_no", data.get("phone", ""))).strip()
    customer_name = mandal_name
    items = data.get("items", [])

    if not mandal_name:
        return jsonify({"success": False, "message": "Please enter Mandal Name."}), 400
    if not firm_name:
        return jsonify({"success": False, "message": "Please enter Firm Name."}), 400
    if not re.fullmatch(r"[0-9]{10}", phone):
        return jsonify({"success": False, "message": "Please enter a valid 10-digit mobile number."}), 400
    if not isinstance(items, list) or not items:
        return jsonify({"success": False, "message": "Please select at least one product."}), 400

    product_by_id = {str(p["id"]): p for p in PRODUCTS}
    order_items, total = [], 0
    for item in items:
        if not isinstance(item, dict):
            continue
        product = product_by_id.get(str(item.get("product_id", item.get("id", ""))))
        try:
            quantity = int(item.get("quantity", 0))
        except (TypeError, ValueError):
            quantity = 0
        if not product or quantity < 1 or quantity > 100:
            continue
        line_total = int(product["price"]) * quantity
        total += line_total
        order_items.append({
            "code": product["code"], "product_id": product["id"], "name": product["name"],
            "category": product["category"], "unit": product["unit"], "quantity": quantity,
            "unit_price": int(product["price"]), "line_total": line_total
        })

    if not order_items:
        return jsonify({"success": False, "message": "No valid products were selected."}), 400

    now = datetime.now().astimezone()
    order_number = reserve_order_number()
    username = str(session.get("username", "unknown"))
    order = {
        "order_id": str(order_number),
        "username": username,
        "created_at": now.isoformat(timespec="seconds"),
        "customer_name": customer_name,
        "mandal_name": mandal_name,
        "firm_name": firm_name,
        "phone": phone,
        "mobile_no": phone,
        "items": order_items,
        "total": total,
        "status": "Order request received — confirm stock and final price with the customer."
    }
    pdf_bytes = build_order_pdf(order)
    pdf_filename = f'{order["order_id"]}_order.pdf'
    # Keep each user's order PDFs in their own directory on the persistent volume.
    user_folder_name = re.sub(r"[^a-zA-Z0-9_-]", "_", username) or "unknown"
    user_orders_dir = ORDERS_DIR / user_folder_name
    user_orders_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = user_orders_dir / pdf_filename

    try:
        # Store both PDF and Excel receipt in the user's folder and a central all-orders folder.
        pdf_path.write_bytes(pdf_bytes)
        excel_filename = f'{order["order_id"]}_order.xlsx'
        excel_path = user_orders_dir / excel_filename
        build_order_excel(order, excel_path)
        all_pdf_path = ALL_ORDERS_DIR / pdf_filename
        all_excel_path = ALL_ORDERS_DIR / excel_filename
        shutil.copy2(pdf_path, all_pdf_path)
        shutil.copy2(excel_path, all_excel_path)
        # Store order history in SQLite on the same persistent Volume (no order.json files).
        with db_connection() as db:
            db.execute("""
                INSERT INTO orders (order_id, created_at, customer_name, phone, total, status, pdf_filename, order_details, username)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                order["order_id"], order["created_at"], customer_name, phone, total,
                order["status"], f"{user_folder_name}/{pdf_filename}", json.dumps(order, ensure_ascii=False), username
            ))
        rebuild_all_orders_workbook()
    except Exception:
        # Do not report success if durable storage failed.
        for cleanup_path in (pdf_path, user_orders_dir / f'{order_number}_order.xlsx', ALL_ORDERS_DIR / f'{order_number}_order.pdf', ALL_ORDERS_DIR / f'{order_number}_order.xlsx'):
            try:
                if cleanup_path.exists(): cleanup_path.unlink()
            except OSError:
                pass
        app.logger.exception("Failed to persist order")
        return jsonify({"success": False, "message": "Could not save the order. Please try again or contact the stall."}), 500

    import base64
    return jsonify({
        "success": True,
        "message": "Order saved successfully. Your PDF receipt is ready to download.",
        "order": order,
        "pdf_filename": pdf_filename,
        "excel_filename": excel_filename,
        "excel_url": url_for("download_my_order_excel", order_id=order["order_id"]),
        "pdf_base64": base64.b64encode(pdf_bytes).decode("ascii"),
        "saved_pdf_path": f"orders/{user_folder_name}/{pdf_filename}"
    }), 201


@app.get("/my-orders")
@login_required
def my_orders():
    # Admin can see every account's orders; mandal users see only their own.
    username = str(session.get("username", ""))
    is_admin = session.get("role") == "admin" or username == "admin"
    with db_connection() as db:
        if is_admin:
            rows = db.execute(
                "SELECT order_id, created_at, customer_name, phone, total, status, pdf_filename, order_details, username "
                "FROM orders ORDER BY CAST(order_id AS INTEGER) DESC"
            ).fetchall()
        else:
            rows = db.execute(
                "SELECT order_id, created_at, customer_name, phone, total, status, pdf_filename, order_details, username "
                "FROM orders WHERE username=? ORDER BY CAST(order_id AS INTEGER) DESC", (username,)
            ).fetchall()
    orders = []
    for row in rows:
        item = dict(row)
        try:
            parsed = json.loads(item.get("order_details") or "{}")
            if not isinstance(parsed, dict): parsed = {}
        except (TypeError, json.JSONDecodeError):
            parsed = {}
        parsed.setdefault("items", [])
        item["order_details"] = parsed
        orders.append(item)
    return render_template("my_orders.html", orders=orders, username=username, is_admin=is_admin)


@app.get("/my-orders/<order_id>/pdf")
@login_required
def download_my_order_pdf(order_id):
    username = str(session.get("username", ""))
    is_admin = session.get("role") == "admin" or username == "admin"
    with db_connection() as db:
        if is_admin:
            row = db.execute("SELECT pdf_filename FROM orders WHERE order_id=?", (str(order_id),)).fetchone()
        else:
            row = db.execute("SELECT pdf_filename FROM orders WHERE order_id=? AND username=?", (str(order_id), username)).fetchone()
    if not row:
        abort(404)
    path = ORDERS_DIR / row["pdf_filename"]
    if not path.is_file():
        abort(404)
    return send_file(path, as_attachment=True, download_name=f"{order_id}_order.pdf", mimetype="application/pdf")


@app.get("/my-orders/<order_id>/excel")
@login_required
def download_my_order_excel(order_id):
    username = str(session.get("username", ""))
    is_admin = session.get("role") == "admin" or username == "admin"
    with db_connection() as db:
        if is_admin:
            row = db.execute("SELECT order_details, username FROM orders WHERE order_id=?", (str(order_id),)).fetchone()
        else:
            row = db.execute("SELECT order_details, username FROM orders WHERE order_id=? AND username=?", (str(order_id), username)).fetchone()
    if not row:
        abort(404)
    order = json.loads(row["order_details"])
    order_owner = row["username"] if "username" in row.keys() else username
    folder = re.sub(r"[^a-zA-Z0-9_-]", "_", order_owner or "unknown") or "unknown"
    path = ORDERS_DIR / folder / f"{order_id}_order.xlsx"
    if not path.is_file():
        # Backfill Excel receipts for orders created before this feature.
        build_order_excel(order, path)
    return send_file(path, as_attachment=True, download_name=f"{order_id}_order.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/admin/all-orders.xlsx")
@admin_required
def download_all_orders_excel():
    path = rebuild_all_orders_workbook()
    return send_file(path, as_attachment=True, download_name="all_orders.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/admin/user-order-stock.xlsx")
@admin_required
def download_user_order_stock_excel():
    """Admin-only product matrix: quantities ordered by each account plus current stock."""
    customer_users = [username for username, info in LOGIN_USERS.items() if info.get("role") == "customer"]
    with db_connection() as db:
        stock_rows = db.execute("SELECT product_id, stock FROM inventory").fetchall()
        order_rows = db.execute("SELECT username, order_details FROM orders").fetchall()

    stock_by_id = {str(row["product_id"]): int(row["stock"]) for row in stock_rows}
    quantities = {username: {} for username in customer_users}
    for row in order_rows:
        username = row["username"] or "unknown"
        if username not in quantities:
            quantities[username] = {}
            customer_users.append(username)
        try:
            details = json.loads(row["order_details"] or "{}")
        except (TypeError, json.JSONDecodeError):
            continue
        for item in details.get("items", []):
            code = str(item.get("code", ""))
            try:
                qty = int(item.get("quantity", 0))
            except (TypeError, ValueError):
                qty = 0
            quantities[username][code] = quantities[username].get(code, 0) + qty

    wb = Workbook()
    ws = wb.active
    ws.title = "User Orders and Stock"
    headers = ["No.", "Item", "Product Code"] + customer_users + ["Total Ordered", "Remaining Stock"]
    ws.append(headers)
    for number, product in enumerate(PRODUCTS, start=1):
        code = str(product.get("code", ""))
        ordered_by_user = [quantities.get(username, {}).get(code, 0) for username in customer_users]
        total_ordered = sum(ordered_by_user)
        remaining = stock_by_id.get(str(product.get("id")), 100)
        ws.append([number, product.get("name", ""), code] + ordered_by_user + [total_ordered, remaining])
    style_workbook(ws)
    ws.freeze_panes = "D2"
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 16
    for col in range(4, len(headers) + 1):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = 24 if col <= 3 + len(customer_users) else 18
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="kkp_user_orders_remaining_stock.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/admin/stock.xlsx")
@admin_required
def download_stock_excel():
    with db_connection() as db:
        rows = db.execute("SELECT product_id, stock FROM inventory").fetchall()
    stock_by_id = {row["product_id"]: int(row["stock"]) for row in rows}
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock Report"
    ws.append(["Product Code", "Product Name", "Category", "Current Stock", "Units To Replenish"])
    for product in PRODUCTS:
        current = stock_by_id.get(str(product["id"]), 100)
        ws.append([product["code"], product["name"], product["category"], current, max(0, -current)])
    style_workbook(ws)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="kkp_stock_report.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def build_whatsapp_url(phone, message):
    """Build a WhatsApp click-to-chat link from a customer mobile number."""
    digits = re.sub(r"\D", "", str(phone or ""))
    if len(digits) == 10:
        digits = "91" + digits
    if not digits:
        return ""
    return "https://wa.me/" + digits + "?text=" + quote_plus(message)


def attach_delivery_links(order):
    details = order.get("order_details", {})
    items = details.get("items", []) if isinstance(details, dict) else []
    lines = [
        f"Hello {order.get('customer_name') or 'Customer'},",
        f"This is regarding your KKP Fireworks order #{order.get('order_id')}.",
        "We are contacting you about delivery/collection of your order.",
        "Please reply to confirm a convenient time.",
        "Thank you, KKP Fireworks."
    ]
    order["whatsapp_url"] = build_whatsapp_url(order.get("phone"), "\n".join(lines))
    delivery_lines = [
        f"Hello {order.get('customer_name') or 'Customer'},",
        f"Your KKP Fireworks order #{order.get('order_id')} is being arranged for delivery/collection.",
        "Items: " + (", ".join(f"{i.get('name', i.get('code', 'Item'))} x {i.get('quantity', 0)}" for i in items) or "As per your order"),
        f"Order total: Rs. {int(order.get('total') or 0):,}",
        "Please reply to coordinate delivery. Thank you, KKP Fireworks."
    ]
    order["delivery_whatsapp_url"] = build_whatsapp_url(order.get("phone"), "\n".join(delivery_lines))
    return order


@app.get("/admin")
@admin_required
def admin_dashboard():
    with db_connection() as db:
        rows = db.execute("SELECT order_id, created_at, customer_name, phone, total, status, order_details, username FROM orders ORDER BY CAST(order_id AS INTEGER) DESC").fetchall()
        stock_rows = db.execute("SELECT product_id, stock FROM inventory").fetchall()
    stock = {row["product_id"]: row["stock"] for row in stock_rows}
    inventory = [dict(p, stock=stock.get(str(p["id"]), 100)) for p in PRODUCTS]
    orders = []
    for row in rows:
        item = dict(row)
        # A malformed historical order must not crash the entire admin dashboard.
        try:
            parsed = json.loads(item.get("order_details") or "{}")
            if not isinstance(parsed, dict):
                parsed = {}
        except (TypeError, json.JSONDecodeError):
            parsed = {}
        parsed.setdefault("items", [])
        item["order_details"] = parsed
        item.setdefault("username", "unknown")
        attach_delivery_links(item)
        orders.append(item)
    return render_template("admin.html", inventory=inventory, orders=orders)


@app.get("/admin/delivery-report")
@admin_required
def admin_delivery_report():
    """Admin-only delivery report; reads existing orders without modifying them."""
    with db_connection() as db:
        rows = db.execute(
            "SELECT order_id, created_at, customer_name, phone, total, status, order_details, username "
            "FROM orders ORDER BY CAST(order_id AS INTEGER) DESC"
        ).fetchall()
    orders = []
    for row in rows:
        item = dict(row)
        try:
            parsed = json.loads(item.get("order_details") or "{}")
            if not isinstance(parsed, dict): parsed = {}
        except (TypeError, json.JSONDecodeError):
            parsed = {}
        parsed.setdefault("items", [])
        item["order_details"] = parsed
        item.setdefault("username", "unknown")
        attach_delivery_links(item)
        orders.append(item)
    return render_template("delivery_report.html", orders=orders)


@app.post("/api/admin/stock")
@admin_required
def update_stock():
    data = request.get_json(silent=True) or {}
    product_id = str(data.get("product_id", ""))
    try:
        stock = int(data.get("stock"))
    except (TypeError, ValueError):
        return jsonify({"success": False, "message": "Stock must be a whole number."}), 400
    if stock < -1000000 or stock > 1000000:
        return jsonify({"success": False, "message": "Stock must be between -1,000,000 and 1,000,000."}), 400
    if product_id not in {str(p["id"]) for p in PRODUCTS}:
        return jsonify({"success": False, "message": "Unknown product."}), 404
    with db_connection() as db:
        db.execute("UPDATE inventory SET stock=?, updated_at=? WHERE product_id=?",
                   (stock, datetime.now().astimezone().isoformat(timespec="seconds"), product_id))
    return jsonify({"success": True, "message": "Stock updated.", "stock": stock})


@app.post("/api/admin/orders/<order_id>/confirm")
@admin_required
def confirm_order(order_id):
    # Deduct stock only when the admin confirms the order. Transaction prevents overselling.
    with db_connection() as db:
        db.execute("BEGIN IMMEDIATE")
        row = db.execute("SELECT status, order_details FROM orders WHERE order_id=?", (str(order_id),)).fetchone()
        if not row:
            return jsonify({"success": False, "message": "Order not found."}), 404
        if row["status"] == "Confirmed":
            return jsonify({"success": False, "message": "This order is already confirmed."}), 409
        order = json.loads(row["order_details"])
        # Orders may exceed current stock. Confirming deducts the full quantity,
        # allowing negative stock to show the amount the admin needs to replenish.
        for item in order["items"]:
            db.execute("UPDATE inventory SET stock=stock-?, updated_at=? WHERE product_id=?",
                       (int(item["quantity"]), datetime.now().astimezone().isoformat(timespec="seconds"), str(item["product_id"])))
        order["status"] = "Confirmed"
        db.execute("UPDATE orders SET status=?, order_details=? WHERE order_id=?",
                   ("Confirmed", json.dumps(order, ensure_ascii=False), str(order_id)))
        db.commit()
    return jsonify({"success": True, "message": "Order confirmed and stock deducted."})


@app.get("/health")
def health():
    return jsonify({"status": "ok", "products": len(PRODUCTS)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)

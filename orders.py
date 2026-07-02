from flask import Flask, request, Response
import win32print
import sqlite3
import json
import os
import shutil
import threading
import time
from datetime import datetime, timedelta

app = Flask(__name__)

# ── CONFIG ──────────────────────────────────────────────
PRINTER_NAME = "POS80 (1)"
DASHBOARD_PASSWORD = "0000"  # change this to whatever you like
DB_PATH = "orders.db"
SETTINGS_PATH = "settings.json"
# ────────────────────────────────────────────────────────

# ── SETTINGS (special hours + announcement banner) ───────
DEFAULT_SETTINGS = {
    "special_dates": {
        # "2026-12-31": {"open": 11, "close": 24, "label": "New Year's Eve — Open till Midnight!"}
        # "2026-08-15": {"open": None, "close": None, "label": "Closed for Independence Day"}
    },
    "banner": {
        "active": False,
        "text": "",
        "emoji": "🎉"
    }
}


def load_settings():
    if not os.path.exists(SETTINGS_PATH):
        save_settings(DEFAULT_SETTINGS)
        return DEFAULT_SETTINGS
    try:
        with open(SETTINGS_PATH, "r") as f:
            data = json.load(f)
            # Backfill any missing keys (e.g. after an app update)
            merged = {**DEFAULT_SETTINGS, **data}
            merged.setdefault("special_dates", {})
            merged.setdefault("banner", DEFAULT_SETTINGS["banner"])
            return merged
    except Exception as e:
        print(f"[SETTINGS LOAD ERROR] {e}")
        return DEFAULT_SETTINGS


def save_settings(settings):
    with open(SETTINGS_PATH, "w") as f:
        json.dump(settings, f, indent=2)


# ── ESC/POS COMMANDS ────────────────────────────────────
ESC = b'\x1b'
GS  = b'\x1d'
INIT         = ESC + b'@'
ALIGN_CENTER = ESC + b'a\x01'
ALIGN_LEFT   = ESC + b'a\x00'
BOLD_ON      = ESC + b'E\x01'
BOLD_OFF     = ESC + b'E\x00'
DOUBLE_ON    = GS  + b'!\x11'
DOUBLE_OFF   = GS  + b'!\x00'
CUT          = GS  + b'V\x41\x03'
LF           = b'\n'
# ────────────────────────────────────────────────────────


# ── DATABASE SETUP ──────────────────────────────────────
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_number TEXT,
            customer_name TEXT,
            customer_phone TEXT,
            order_type TEXT,
            address TEXT,
            notes TEXT,
            items_json TEXT,
            subtotal INTEGER,
            packing_charge INTEGER,
            delivery_charge INTEGER,
            grand_total INTEGER,
            order_time TEXT,
            status TEXT DEFAULT 'pending',
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()


init_db()


# ── AUTOMATED BACKUPS ────────────────────────────────────
# Copies orders.db into a dated file inside BACKUP_DIR once a day, and
# prunes anything older than BACKUP_RETENTION_DAYS. This protects against
# disk failure / accidental deletion on the machine running this script.
#
# IMPORTANT: for real protection against this PC failing entirely, point
# BACKUP_DIR at a folder that syncs off this machine — e.g. install the
# Google Drive desktop app and set BACKUP_DIR to a path inside your
# Google Drive folder. Otherwise these backups live on the same disk as
# the original database and won't survive a hardware failure.
BACKUP_DIR = "backups"
BACKUP_RETENTION_DAYS = 30


def run_backup():
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        today_str = datetime.now().strftime("%Y-%m-%d")
        dest = os.path.join(BACKUP_DIR, f"orders_backup_{today_str}.db")
        if not os.path.exists(dest):
            shutil.copy2(DB_PATH, dest)
            print(f"[BACKUP] Created {dest}")
        # Prune backups older than the retention window
        cutoff = datetime.now() - timedelta(days=BACKUP_RETENTION_DAYS)
        for fname in os.listdir(BACKUP_DIR):
            fpath = os.path.join(BACKUP_DIR, fname)
            if os.path.isfile(fpath) and datetime.fromtimestamp(os.path.getmtime(fpath)) < cutoff:
                os.remove(fpath)
                print(f"[BACKUP] Pruned old backup {fname}")
    except Exception as e:
        print(f"[BACKUP ERROR] {e}")


def backup_loop():
    while True:
        run_backup()
        time.sleep(6 * 60 * 60)  # check every 6 hours; run_backup() is a no-op if today's backup already exists


threading.Thread(target=backup_loop, daemon=True).start()
# ────────────────────────────────────────────────────────


def save_order(data):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Idempotency guard: if the browser retries a submission (e.g. after a
    # dropped connection where the request actually succeeded but the
    # response never made it back), don't create a duplicate kitchen order.
    # Same order_number + phone within the last 30 minutes = same order.
    order_number = data.get("order_number", "")
    customer_phone = data.get("customer_phone", "")
    c.execute("""
        SELECT id FROM orders
        WHERE order_number = ? AND customer_phone = ?
          AND created_at >= datetime('now', '-30 minutes')
        LIMIT 1
    """, (order_number, customer_phone))
    existing = c.fetchone()
    if existing:
        conn.close()
        return existing[0]

    c.execute("""
        INSERT INTO orders (order_number, customer_name, customer_phone, order_type,
            address, notes, items_json, subtotal, packing_charge, delivery_charge,
            grand_total, order_time, status, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?)
    """, (
        order_number,
        data.get("customer_name", ""),
        customer_phone,
        data.get("order_type", ""),
        data.get("address", ""),
        data.get("notes", ""),
        json.dumps(data.get("items", [])),
        data.get("subtotal", 0),
        data.get("packing_charge", 0),
        data.get("delivery_charge", 0),
        data.get("grand_total", 0),
        data.get("order_time", ""),
        datetime.now().isoformat()
    ))
    conn.commit()
    order_id = c.lastrowid
    conn.close()
    return order_id


def get_pending_orders():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM orders WHERE status = 'pending' ORDER BY id DESC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_orders_by_date(date_str):
    """All orders (any status) placed on a single given date, newest first."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM orders WHERE date(created_at) = date(?) ORDER BY id DESC", (date_str,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_orders_in_range(start_date, end_date):
    """All orders (any status) between two dates, inclusive — used for CSV export."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("""
        SELECT * FROM orders
        WHERE date(created_at) BETWEEN date(?) AND date(?)
        ORDER BY created_at
    """, (start_date, end_date))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_order_by_id(order_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM orders WHERE id = ?", (order_id,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def mark_order_status(order_id, status):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE orders SET status = ? WHERE id = ?", (status, order_id))
    conn.commit()
    conn.close()


def confirm_order_db(order_id, delivery_charge, grand_total, status):
    """Persist the manually-entered delivery charge and the recalculated
    grand total alongside the new status. This is what api_confirm_order
    should call so that confirmed orders are stored with their real totals
    (previously the delivery charge only lived in the printed receipt and
    was never saved back to the database)."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "UPDATE orders SET delivery_charge = ?, grand_total = ?, status = ? WHERE id = ?",
        (delivery_charge, grand_total, status, order_id)
    )
    conn.commit()
    conn.close()


def search_orders(query, limit=100, date_str=None):
    """Search orders by order number or phone number, newest first.
    If date_str is given, results are restricted to that single date
    (used for e.g. 'find today's order from this phone number');
    otherwise it searches across all dates."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    like_query = "%" + query.strip() + "%"
    if date_str:
        c.execute("""
            SELECT * FROM orders
            WHERE (order_number LIKE ? OR customer_phone LIKE ?)
              AND date(created_at) = date(?)
            ORDER BY id DESC
            LIMIT ?
        """, (like_query, like_query, date_str, limit))
    else:
        c.execute("""
            SELECT * FROM orders
            WHERE order_number LIKE ? OR customer_phone LIKE ?
            ORDER BY id DESC
            LIMIT ?
        """, (like_query, like_query, limit))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_recent_orders(limit=50):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM orders ORDER BY id DESC LIMIT ?", (limit,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_sales_summary(start_date, end_date):
    """Aggregate CONFIRMED (status='printed') orders between start_date and
    end_date (inclusive, 'YYYY-MM-DD' strings). Uses created_at (a real
    ISO timestamp written by the server) rather than order_time (a
    free-text string from the browser) since it sorts/filters reliably."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("""
        SELECT * FROM orders
        WHERE status = 'printed'
          AND date(created_at) BETWEEN date(?) AND date(?)
        ORDER BY created_at
    """, (start_date, end_date))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()

    total_revenue = 0
    by_type = {}
    by_day = {}
    item_stats = {}

    for o in rows:
        gt = o.get("grand_total") or 0
        total_revenue += gt

        otype = o.get("order_type") or "Unknown"
        t = by_type.setdefault(otype, {"count": 0, "revenue": 0})
        t["count"] += 1
        t["revenue"] += gt

        day = (o.get("created_at") or "")[:10]
        d = by_day.setdefault(day, {"count": 0, "revenue": 0})
        d["count"] += 1
        d["revenue"] += gt

        try:
            items = json.loads(o.get("items_json") or "[]")
        except Exception:
            items = []
        for it in items:
            name = it.get("product_retailer_id", "Unknown")
            qty = it.get("quantity", 0) or 0
            price = it.get("item_price", 0) or 0
            entry = item_stats.setdefault(name, {"qty": 0, "revenue": 0})
            entry["qty"] += qty
            entry["revenue"] += qty * price

    by_type_list = [{"type": k, "count": v["count"], "revenue": v["revenue"]} for k, v in by_type.items()]
    by_day_list = sorted(
        [{"day": k, "count": v["count"], "revenue": v["revenue"]} for k, v in by_day.items()],
        key=lambda x: x["day"]
    )
    top_items = sorted(
        [{"name": k, "qty": v["qty"], "revenue": v["revenue"]} for k, v in item_stats.items()],
        key=lambda x: x["qty"], reverse=True
    )[:10]

    return {
        "start_date": start_date,
        "end_date": end_date,
        "total_orders": len(rows),
        "total_revenue": total_revenue,
        "by_type": by_type_list,
        "by_day": by_day_list,
        "top_items": top_items
    }


# ── PRINTING ─────────────────────────────────────────────
def raw_print(data: bytes):
    try:
        hPrinter = win32print.OpenPrinter(PRINTER_NAME)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Order", None, "RAW"))
            try:
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, data)
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
        print("[PRINT] Job sent successfully")
        return True
    except Exception as e:
        print(f"[PRINT ERROR] {e}")
        return False


def encode(text):
    return text.encode('ascii', errors='replace') + LF


def divider(char="-", width=48):
    return char * width


def print_order(order):
    customer_name = order["customer_name"]
    customer_phone = order["customer_phone"]
    items = json.loads(order["items_json"]) if isinstance(order["items_json"], str) else order["items_json"]
    order_number = order["order_number"]
    order_time = order["order_time"]
    order_type = order["order_type"]
    address = order.get("address", "") or ""
    notes = order.get("notes", "") or ""
    subtotal = order.get("subtotal", 0)
    packing_charge = order.get("packing_charge", 0)
    delivery_charge = order.get("delivery_charge", 0)

    buf = bytearray()

    def add(data: bytes):
        buf.extend(data)

    # ═══════════════════════════════
    # RECEIPT
    # ═══════════════════════════════
    add(INIT)
    add(ALIGN_CENTER)
    add(BOLD_ON)
    add(DOUBLE_ON)
    add(encode("Grill Inn"))
    add(DOUBLE_OFF)
    add(BOLD_OFF)
    add(encode("Falkawn"))
    add(encode("9612992023"))
    add(encode("GSTIN: 15ARKPVB080N1ZX"))
    add(LF)
    add(BOLD_ON)
    add(encode(order_type))
    add(BOLD_OFF)
    add(LF)

    add(ALIGN_LEFT)
    add(encode(f"Order #: {order_number}"))
    add(encode(f"Date   : {order_time}"))
    add(LF)
    add(encode(f"Customer: {customer_name}"))
    add(encode(f"Phone   : {customer_phone}"))
    if order_type == "Delivery" and address:
        add(encode(f"Address : {address}"))
    add(LF)

    add(encode(divider()))
    add(encode(f"{'Item':<24}{'Qty':>6}{'Rate':>8}{'Amt':>8}"))
    add(encode(divider()))

    for item in items:
        name = item.get("product_retailer_id", "Unknown")[:23]
        qty = item.get("quantity", 1)
        price = item.get("item_price", 0)
        amt = qty * price
        add(encode(f"{name:<24}{qty:>6}{price:>8}{amt:>8}"))

    add(encode(divider()))
    add(encode(f"{'Sub Total':<32}{subtotal:>16}"))
    if packing_charge > 0:
        add(encode(f"{'Packing Charges':<32}{packing_charge:>16}"))
    if order_type == "Delivery":
        dc_str = str(delivery_charge) if delivery_charge else "______"
        add(encode(f"{'Delivery Charges':<32}{dc_str:>16}"))
    add(encode(divider()))
    grand = subtotal + packing_charge + (delivery_charge if order_type == "Delivery" else 0)
    add(encode(f"{'Bill Total':<32}{grand:>16}"))
    add(LF)
    add(encode("Payment: Cash on Delivery"))

    if notes:
        add(LF)
        add(encode(f"Notes: {notes[:40]}"))

    add(LF)
    add(ALIGN_CENTER)
    add(BOLD_ON)
    add(encode("!! Thanks For Ordering !!"))
    add(BOLD_OFF)
    add(LF)
    add(LF)
    add(CUT)

    # ═══════════════════════════════
    # KOT
    # ═══════════════════════════════
    add(INIT)
    add(ALIGN_CENTER)
    add(BOLD_ON)
    add(DOUBLE_ON)
    add(encode("Grill Inn"))
    add(DOUBLE_OFF)
    add(BOLD_OFF)
    add(encode("Service Ticket: KOT"))
    add(LF)
    add(BOLD_ON)
    add(encode(order_type))
    add(BOLD_OFF)
    add(LF)

    add(ALIGN_LEFT)
    add(encode(f"Order #: {order_number}"))
    add(encode(f"Date   : {order_time}"))
    add(LF)
    add(encode(f"Customer: {customer_name}"))
    add(encode(f"Phone   : {customer_phone}"))
    if order_type == "Delivery" and address:
        add(encode(f"Address : {address}"))
    if notes:
        add(encode(f"Notes  : {notes[:40]}"))
    add(LF)

    add(encode(divider()))
    add(encode(f"{'Item':<36}{'Qty':>12}"))
    add(encode(divider()))

    for item in items:
        name = item.get("product_retailer_id", "Unknown")[:35]
        qty = item.get("quantity", 1)
        add(encode(f"{name:<36}{qty:>12}"))

    add(encode(divider()))
    add(LF)
    add(ALIGN_CENTER)
    add(BOLD_ON)
    add(encode("** ONLINE ORDER **"))
    add(BOLD_OFF)
    add(LF)
    add(LF)
    add(LF)
    add(CUT)

    success = raw_print(bytes(buf))
    print(f"[PRINTED] Order #{order_number} for {customer_name} — {order_type} — Success: {success}")
    return success


# ── WEBSITE ORDER ROUTE ──────────────────────────────────
@app.route("/web-order", methods=["POST", "OPTIONS"])
def web_order():
    if request.method == "OPTIONS":
        response = Response("", status=200)
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    data = request.get_json()
    print(f"[WEB ORDER] {data}")

    try:
        order_id = save_order(data)
        print(f"[ORDER SAVED] DB ID #{order_id} — Order #{data.get('order_number')} — PENDING CONFIRMATION")

        response = Response(
            json.dumps({"status": "ok", "order_id": order_id}),
            status=200,
            mimetype="application/json"
        )
        response.headers["Access-Control-Allow-Origin"] = "*"
        return response

    except Exception as e:
        print(f"[WEB ORDER ERROR] {e}")
        response = Response('{"status":"error"}', status=500, mimetype="application/json")
        response.headers["Access-Control-Allow-Origin"] = "*"
        return response


# ── SETTINGS ROUTES (special hours + banner) ──────────────
@app.route("/api/settings", methods=["GET", "OPTIONS"])
def api_get_settings():
    if request.method == "OPTIONS":
        response = Response("", status=200)
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "GET, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    settings = load_settings()
    response = Response(json.dumps(settings), status=200, mimetype="application/json")
    response.headers["Access-Control-Allow-Origin"] = "*"
    return response


@app.route("/api/settings", methods=["POST"])
def api_save_settings():
    data = request.get_json() or {}

    if data.get("password") != DASHBOARD_PASSWORD:
        return Response('{"status":"error","message":"Incorrect password"}', status=401, mimetype="application/json")

    settings = load_settings()
    if "special_dates" in data:
        settings["special_dates"] = data["special_dates"]
    if "banner" in data:
        settings["banner"] = data["banner"]

    save_settings(settings)
    return Response(json.dumps({"status": "ok"}), status=200, mimetype="application/json")


# ── DASHBOARD ROUTES ──────────────────────────────────────
@app.route("/api/download-backup")
def api_download_backup():
    password = request.args.get("password", "")
    if password != DASHBOARD_PASSWORD:
        return Response('{"status":"error","message":"Incorrect password"}', status=401, mimetype="application/json")

    run_backup()  # make sure today's backup exists before serving it
    try:
        with open(DB_PATH, "rb") as f:
            data = f.read()
        today_str = datetime.now().strftime("%Y-%m-%d")
        response = Response(data, status=200, mimetype="application/octet-stream")
        response.headers["Content-Disposition"] = f'attachment; filename="grillinn_orders_backup_{today_str}.db"'
        return response
    except Exception as e:
        return Response(json.dumps({"status": "error", "message": str(e)}), status=500, mimetype="application/json")


@app.route("/dashboard")
def dashboard():
    return DASHBOARD_HTML


@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json()
    if data.get("password") == DASHBOARD_PASSWORD:
        return Response('{"status":"ok"}', status=200, mimetype="application/json")
    return Response('{"status":"error"}', status=401, mimetype="application/json")


@app.route("/api/pending-orders")
def api_pending_orders():
    orders = get_pending_orders()
    return Response(json.dumps(orders), status=200, mimetype="application/json")


@app.route("/api/recent-orders")
def api_recent_orders():
    date_param = request.args.get("date")
    if date_param:
        orders = get_orders_by_date(date_param)
    else:
        orders = get_recent_orders(100)
    return Response(json.dumps(orders), status=200, mimetype="application/json")


@app.route("/api/search-orders")
def api_search_orders():
    query = request.args.get("q", "").strip()
    date_param = request.args.get("date", "").strip() or None
    if not query:
        return Response(json.dumps([]), status=200, mimetype="application/json")
    orders = search_orders(query, date_str=date_param)
    return Response(json.dumps(orders), status=200, mimetype="application/json")


@app.route("/api/confirm-order/<int:order_id>", methods=["POST"])
def api_confirm_order(order_id):
    order = get_order_by_id(order_id)
    if not order:
        return Response('{"status":"not_found"}', status=404, mimetype="application/json")

    # Get manually entered delivery charge from dashboard
    data = request.get_json() or {}
    manual_delivery_charge = int(data.get("delivery_charge", 0))
    delivery_charge = manual_delivery_charge if order.get("order_type") == "Delivery" else 0
    order["delivery_charge"] = delivery_charge

    subtotal = order.get("subtotal", 0) or 0
    packing_charge = order.get("packing_charge", 0) or 0
    grand_total = subtotal + packing_charge + delivery_charge

    success = print_order(order)
    # Persist the confirmed delivery charge + recalculated grand total,
    # not just the status — otherwise sales reports would undercount
    # every Delivery order by its delivery charge.
    confirm_order_db(order_id, delivery_charge, grand_total, "printed" if success else "print_failed")

    return Response(
        json.dumps({"status": "ok" if success else "print_failed"}),
        status=200,
        mimetype="application/json"
    )


@app.route("/api/sales-summary")
def api_sales_summary():
    today = datetime.now().strftime("%Y-%m-%d")
    start = request.args.get("start") or today
    end = request.args.get("end") or today
    summary = get_sales_summary(start, end)
    return Response(json.dumps(summary), status=200, mimetype="application/json")


@app.route("/api/export-orders-csv")
def api_export_orders_csv():
    import io
    import csv as csv_module

    today = datetime.now().strftime("%Y-%m-%d")
    start = request.args.get("start") or today
    end = request.args.get("end") or today
    orders = get_orders_in_range(start, end)

    output = io.StringIO()
    writer = csv_module.writer(output)
    writer.writerow([
        "Order #", "Date/Time", "Order Type", "Customer", "Phone", "Address",
        "Items", "Subtotal", "Packing Charge", "Delivery Charge", "Grand Total",
        "Status", "Notes"
    ])
    for o in orders:
        try:
            items = json.loads(o.get("items_json") or "[]")
        except Exception:
            items = []
        items_str = "; ".join(
            f'{it.get("quantity", 0)}x {it.get("product_retailer_id", "")}' for it in items
        )
        writer.writerow([
            o.get("order_number", ""),
            o.get("created_at", ""),
            o.get("order_type", ""),
            o.get("customer_name", ""),
            o.get("customer_phone", ""),
            o.get("address", ""),
            items_str,
            o.get("subtotal", 0),
            o.get("packing_charge", 0),
            o.get("delivery_charge", 0),
            o.get("grand_total", 0),
            o.get("status", ""),
            o.get("notes", "")
        ])

    response = Response(output.getvalue(), status=200, mimetype="text/csv")
    response.headers["Content-Disposition"] = f"attachment; filename=orders_{start}_to_{end}.csv"
    return response


@app.route("/api/reject-order/<int:order_id>", methods=["POST"])
def api_reject_order(order_id):
    mark_order_status(order_id, "rejected")
    return Response('{"status":"ok"}', status=200, mimetype="application/json")


@app.route("/api/cancel-order/<int:order_id>", methods=["POST"])
def api_cancel_order(order_id):
    order = get_order_by_id(order_id)
    if not order:
        return Response('{"status":"not_found"}', status=404, mimetype="application/json")
    mark_order_status(order_id, "cancelled")
    return Response('{"status":"ok"}', status=200, mimetype="application/json")


@app.route("/api/reprint-order/<int:order_id>", methods=["POST"])
def api_reprint_order(order_id):
    order = get_order_by_id(order_id)
    if not order:
        return Response('{"status":"not_found"}', status=404, mimetype="application/json")
    success = print_order(order)
    return Response(
        json.dumps({"status": "ok" if success else "print_failed"}),
        status=200,
        mimetype="application/json"
    )


# ── DASHBOARD HTML (served at /dashboard) ─────────────────
DASHBOARD_HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Grill Inn - Order Dashboard</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Inter:wght@400;500;600;700&display=swap');
  :root{--black:#0f0f0f;--charcoal:#1a1a1a;--panel:#222;--border:#2e2e2e;--orange:#f97316;--red:#dc2626;--flame:#ea580c;--white:#f5f5f5;--muted:#888;--green:#4ade80;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--black);color:var(--white);font-family:'Inter',sans-serif;min-height:100vh;}
  .login-screen{display:flex;align-items:center;justify-content:center;min-height:100vh;flex-direction:column;gap:1rem;padding:2rem;}
  .login-screen.hidden{display:none;}
  .logo{font-family:'Bebas Neue',sans-serif;font-size:2.5rem;letter-spacing:2px;background:linear-gradient(135deg,var(--orange),var(--red));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;}
  .login-box{background:var(--charcoal);border:1px solid var(--border);border-radius:12px;padding:2rem;width:100%;max-width:320px;}
  .login-box input{width:100%;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--white);font-size:1rem;padding:0.75rem;margin-bottom:1rem;outline:none;}
  .login-box input:focus{border-color:var(--orange);}
  .login-box button{width:100%;background:linear-gradient(135deg,var(--orange),var(--red));border:none;border-radius:8px;color:white;font-weight:700;padding:0.8rem;cursor:pointer;font-size:0.95rem;}
  .app{display:none;}
  .app.visible{display:block;}
  header{background:var(--charcoal);border-bottom:2px solid var(--flame);padding:1rem;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:50;}
  .tabs{display:flex;gap:0.5rem;padding:1rem;background:var(--charcoal);border-bottom:1px solid var(--border);}
  .tab-btn{background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);cursor:pointer;font-weight:700;padding:0.6rem 1.2rem;font-size:0.85rem;}
  .tab-btn.active{background:var(--flame);border-color:var(--flame);color:white;}
  .main{max-width:800px;margin:0 auto;padding:1rem;}
  .order-card{background:var(--charcoal);border:1px solid var(--border);border-radius:12px;padding:1rem;margin-bottom:1rem;}
  .order-card.pending{border-color:var(--orange);box-shadow:0 0 0 1px var(--orange);}
  .order-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:0.5rem;}
  .order-num{font-family:'Bebas Neue',sans-serif;font-size:1.4rem;color:var(--orange);letter-spacing:1px;}
  .order-type-tag{font-size:0.7rem;font-weight:700;padding:0.2rem 0.6rem;border-radius:10px;}
  .order-type-tag.Delivery{background:rgba(249,115,22,0.15);color:var(--orange);}
  .order-type-tag.Takeaway{background:rgba(74,222,128,0.15);color:var(--green);}
  .order-type-tag.Dine-In{background:rgba(96,165,250,0.15);color:#60a5fa;}
  .order-meta{font-size:0.8rem;color:var(--muted);margin-bottom:0.5rem;}
  .order-items{background:var(--panel);border-radius:8px;padding:0.6rem 0.8rem;margin-bottom:0.6rem;font-size:0.85rem;}
  .order-item-row{display:flex;justify-content:space-between;padding:0.2rem 0;}
  .order-total{font-weight:700;color:var(--orange);text-align:right;margin-top:0.4rem;padding-top:0.4rem;border-top:1px solid var(--border);}
  .order-actions{display:flex;gap:0.5rem;}
  .btn-confirm{flex:1;background:linear-gradient(135deg,var(--green),#16a34a);border:none;border-radius:8px;color:white;font-weight:700;padding:0.7rem;cursor:pointer;font-size:0.9rem;}
  .dc-field-label{font-size:0.72rem;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:0.25rem;}
  .dc-input{width:100%;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--white);font-size:0.95rem;padding:0.6rem 0.75rem;outline:none;margin-bottom:0.5rem;font-family:'Inter',sans-serif;-moz-appearance:textfield;}
  .dc-input:focus{border-color:var(--orange);background:#262626;}
  .dc-input::-webkit-outer-spin-button,.dc-input::-webkit-inner-spin-button{-webkit-appearance:none;margin:0;}
  .history-search-row{display:flex;gap:0.5rem;margin:0 0 0.8rem;}
  .history-search-input{flex:1;background:var(--charcoal);border:1px solid var(--border);border-radius:8px;color:var(--white);font-size:0.9rem;padding:0.65rem 0.8rem;outline:none;font-family:'Inter',sans-serif;}
  .history-search-input:focus{border-color:var(--orange);}
  .history-search-clear{background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.5rem 0.9rem;cursor:pointer;font-size:0.85rem;white-space:nowrap;}
  .history-search-scope{display:flex;align-items:center;gap:0.4rem;font-size:0.78rem;color:var(--muted);margin:-0.4rem 0 0.8rem;cursor:pointer;}
  .history-search-scope input{cursor:pointer;}
  .btn-reject{background:var(--panel);border:1px solid var(--red);border-radius:8px;color:var(--red);font-weight:700;padding:0.7rem 1rem;cursor:pointer;font-size:0.9rem;}
  .btn-reprint{background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.5rem 0.9rem;cursor:pointer;font-size:0.8rem;}
  .status-tag{font-size:0.7rem;font-weight:700;padding:0.2rem 0.6rem;border-radius:10px;}
  .status-tag.printed{background:rgba(74,222,128,0.15);color:var(--green);}
  .status-tag.rejected{background:rgba(220,38,38,0.15);color:var(--red);}
  .status-tag.cancelled{background:rgba(220,38,38,0.15);color:var(--red);}
  .status-tag.print_failed{background:rgba(220,38,38,0.15);color:var(--red);}
  .empty-state{text-align:center;padding:3rem 1rem;color:var(--muted);}
  .empty-icon{font-size:3rem;margin-bottom:0.5rem;}
  .badge-count{background:var(--red);color:white;border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;margin-left:0.3rem;}
  .sales-filters{display:flex;gap:0.5rem;flex-wrap:wrap;align-items:center;margin-bottom:1rem;}
  .sales-filters input[type=date]{background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--white);padding:0.5rem 0.6rem;font-size:0.85rem;}
  .sales-quick-btn{background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);cursor:pointer;font-weight:700;padding:0.5rem 0.9rem;font-size:0.8rem;}
  .sales-quick-btn.active{background:var(--flame);border-color:var(--flame);color:white;}
  .sales-summary-card{background:linear-gradient(135deg,var(--orange),var(--red));border-radius:12px;padding:1.2rem;margin-bottom:1rem;text-align:center;}
  .sales-total-label{font-size:0.8rem;color:rgba(255,255,255,0.85);text-transform:uppercase;letter-spacing:0.5px;font-weight:700;}
  .sales-total-value{font-family:'Bebas Neue',sans-serif;font-size:2.4rem;color:white;letter-spacing:1px;}
  .sales-total-sub{font-size:0.8rem;color:rgba(255,255,255,0.85);}
  .sales-section{background:var(--charcoal);border:1px solid var(--border);border-radius:12px;padding:1rem;margin-bottom:1rem;}
  .sales-section-title{font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:var(--muted);margin-bottom:0.6rem;}
  .sales-row{display:flex;justify-content:space-between;gap:0.5rem;padding:0.4rem 0;border-bottom:1px solid var(--border);font-size:0.88rem;}
  .sales-row:last-child{border-bottom:none;}
  .sales-empty{text-align:center;color:var(--muted);padding:2rem 1rem;font-size:0.85rem;}
  .settings-section{background:var(--charcoal);border:1px solid var(--border);border-radius:12px;padding:1rem;margin-bottom:1rem;}
  .settings-section-title{font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:var(--muted);margin-bottom:0.8rem;}
  .field-label{font-size:0.72rem;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:0.25rem;}
  .settings-input{width:100%;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--white);font-size:0.9rem;padding:0.6rem 0.75rem;outline:none;margin-bottom:0.8rem;}
  .settings-input:focus{border-color:var(--orange);}
  .toggle-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:0.8rem;}
  .toggle-switch{position:relative;width:46px;height:26px;background:var(--panel);border:1px solid var(--border);border-radius:13px;cursor:pointer;flex-shrink:0;}
  .toggle-switch.on{background:var(--flame);border-color:var(--flame);}
  .toggle-switch::after{content:'';position:absolute;top:2px;left:2px;width:20px;height:20px;background:white;border-radius:50%;transition:left 0.15s;}
  .toggle-switch.on::after{left:22px;}
  .special-date-card{background:var(--panel);border:1px solid var(--border);border-radius:8px;padding:0.8rem;margin-bottom:0.6rem;}
  .special-date-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:0.5rem;}
  .btn-save{width:100%;background:linear-gradient(135deg,var(--green),#16a34a);border:none;border-radius:8px;color:white;font-weight:700;padding:0.8rem;cursor:pointer;font-size:0.9rem;margin-top:0.4rem;}
  .btn-add-date{width:100%;background:var(--panel);border:1px dashed var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.7rem;cursor:pointer;font-size:0.85rem;margin-bottom:0.6rem;}
  .btn-remove-date{background:none;border:none;color:var(--red);cursor:pointer;font-size:0.9rem;font-weight:700;}
  .row-2col{display:grid;grid-template-columns:1fr 1fr;gap:0.6rem;}
</style>
</head>
<body>

<div class="login-screen" id="loginScreen">
  <div class="logo">Grill Inn</div>
  <div style="color:var(--muted);font-size:0.85rem;margin-bottom:0.5rem;">Order Dashboard</div>
  <div class="login-box">
    <input type="password" id="pwInput" placeholder="Enter password" onkeydown="if(event.key==='Enter')login()">
    <button onclick="login()">Login</button>
  </div>
</div>

<div class="app" id="app">
  <header>
    <div class="logo" style="font-size:1.6rem;">Grill Inn</div>
    <div style="font-size:0.8rem;color:var(--muted);">Order Dashboard</div>
  </header>
  <div class="tabs">
    <button class="tab-btn active" id="pendingTab" onclick="switchTab('pending')">Pending <span class="badge-count" id="pendingCount">0</span></button>
    <button class="tab-btn" id="historyTab" onclick="switchTab('history')">History</button>
    <button class="tab-btn" id="salesTab" onclick="switchTab('sales')">📊 Sales</button>
    <button class="tab-btn" id="settingsTab" onclick="switchTab('settings')">⚙️ Settings</button>
  </div>
  <div class="main" id="mainContent"></div>
</div>

<script>
let currentTab = 'pending';
let knownPendingIds = new Set();
let isFirstLoad = true;
let historyDate = null;
let historySearchQuery = '';
let historySearchAllDates = false;

function login() {
  const pw = document.getElementById('pwInput').value;
  fetch('/api/login', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({password: pw})
  }).then(r => {
    if (r.ok) {
      document.getElementById('loginScreen').classList.add('hidden');
      document.getElementById('app').classList.add('visible');
      sessionStorage.setItem('gi_dash_auth', '1');
      startPolling();
    } else {
      alert('Incorrect password');
    }
  });
}

function switchTab(tab) {
  currentTab = tab;
  document.getElementById('pendingTab').classList.toggle('active', tab === 'pending');
  document.getElementById('historyTab').classList.toggle('active', tab === 'history');
  document.getElementById('salesTab').classList.toggle('active', tab === 'sales');
  document.getElementById('settingsTab').classList.toggle('active', tab === 'settings');
  if (tab === 'sales') {
    loadSales('today');
  } else if (tab === 'settings') {
    loadSettings();
  } else {
    loadOrders();
  }
}

function playNotifySound() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const playBeep = (freq, start, dur) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.frequency.value = freq;
      osc.type = 'square'; // harsher/more alarming than sine
      gain.gain.setValueAtTime(0.0001, ctx.currentTime + start);
      gain.gain.exponentialRampToValueAtTime(0.9, ctx.currentTime + start + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + start + dur);
      osc.start(ctx.currentTime + start);
      osc.stop(ctx.currentTime + start + dur + 0.02);
    };
    // Siren-style pattern, two-tone, repeated 3x — much louder and longer
    // than a single soft chime so it's noticeable from across the kitchen.
    const pattern = [
      [880, 0.00, 0.22], [1320, 0.24, 0.22],
      [880, 0.50, 0.22], [1320, 0.74, 0.22],
      [880, 1.00, 0.22], [1320, 1.24, 0.30],
    ];
    pattern.forEach(([freq, start, dur]) => playBeep(freq, start, dur));
    if (navigator.vibrate) navigator.vibrate([300, 150, 300, 150, 300]);
  } catch(e) { console.log('Sound failed', e); }
}

function fmtItems(itemsJson) {
  let items;
  try { items = JSON.parse(itemsJson); } catch(e) { items = []; }
  return items;
}

function renderOrderCard(order, isPending) {
  const items = fmtItems(order.items_json);
  const itemsHtml = items.map(i =>
    '<div class="order-item-row"><span>' + i.quantity + 'x ' + i.product_retailer_id + '</span><span>₹' + (i.quantity * i.item_price) + '</span></div>'
  ).join('');

  const grandTotal = order.subtotal + (order.packing_charge || 0) + (order.order_type === 'Delivery' ? (order.delivery_charge || 0) : 0);
  const typeClass = order.order_type.replace(' ', '-');

  let actionsHtml = '';
  if (isPending) {
    const dcField = order.order_type === 'Delivery'
      ? '<div style="margin-bottom:0.5rem;"><label class="dc-field-label">Delivery Charge (₹)</label><input id="dc-' + order.id + '" class="dc-input" type="number" inputmode="numeric" min="0" value="0" placeholder="Enter delivery charge..."></div>'
      : '';
    actionsHtml = '<div class="order-actions">' +
      dcField +
      '<div style="display:flex;gap:0.5rem;width:100%;">' +
      '<button class="btn-confirm" style="flex:1;" onclick="confirmOrder(' + order.id + ')">✅ Confirm & Print</button>' +
      '<button class="btn-reject" onclick="rejectOrder(' + order.id + ')">✕</button>' +
      '</div></div>';
  } else {
    const statusLabel = order.status === 'printed' ? '✅ Printed' : order.status === 'rejected' ? '✕ Rejected' : order.status === 'cancelled' ? '🚫 Cancelled' : '⚠️ Print Failed';
    actionsHtml = '<div style="display:flex;justify-content:space-between;align-items:center;gap:0.5rem;flex-wrap:wrap;">' +
      '<span class="status-tag ' + order.status + '">' + statusLabel + '</span>' +
      '<div style="display:flex;gap:0.5rem;">' +
      (order.status === 'printed' ? '<button class="btn-reject" onclick="cancelOrder(' + order.id + ')">🚫 Cancel</button>' : '') +
      '<button class="btn-reprint" onclick="reprintOrder(' + order.id + ')">🖨️ Reprint</button>' +
      '</div></div>';
  }

  return '<div class="order-card ' + (isPending ? 'pending' : '') + '">' +
    '<div class="order-header"><span class="order-num">#' + order.order_number + '</span><span class="order-type-tag ' + typeClass + '">' + order.order_type + '</span></div>' +
    '<div class="order-meta">' + order.customer_name + ' • ' + order.customer_phone + (order.address ? '<br>📍 ' + order.address : '') + '<br>🕒 ' + order.order_time + '</div>' +
    '<div class="order-items">' + itemsHtml + '<div class="order-total">Total: ₹' + grandTotal + '</div></div>' +
    (order.notes ? '<div style="font-size:0.8rem;color:var(--muted);margin-bottom:0.5rem;">📝 ' + order.notes + '</div>' : '') +
    actionsHtml +
    '</div>';
}

// Detects new pending orders and sounds the alarm — runs on its own
// interval regardless of which dashboard tab is currently open (this is
// what used to be missing: previously the check only ran while viewing
// the Pending tab, so switching to History silenced new-order alerts).
// If the Pending tab happens to be open, it also renders the list.
function checkForNewOrders(forceRender) {
  fetch('/api/pending-orders').then(r => r.json()).then(orders => {
    const currentIds = new Set(orders.map(o => o.id));

    if (!isFirstLoad) {
      const newOnes = [...currentIds].filter(id => !knownPendingIds.has(id));
      if (newOnes.length > 0) playNotifySound();
    }

    const idsUnchanged = !isFirstLoad &&
      currentIds.size === knownPendingIds.size &&
      [...currentIds].every(id => knownPendingIds.has(id));

    knownPendingIds = currentIds;
    isFirstLoad = false;
    document.getElementById('pendingCount').textContent = orders.length;

    if (currentTab !== 'pending') return; // don't touch DOM unless it's visible

    // Nothing added or removed since the last poll — skip rebuilding the
    // DOM entirely. This is what used to cause the delivery-charge field
    // to blink/lose focus/reset mid-type: rebuilding the list every 5s
    // recreated the input element even when nothing about the order list
    // had actually changed.
    if (idsUnchanged && !forceRender) return;

    // Preserve whatever the user is actively doing (typing a delivery
    // charge) across the rebuild, since the order list itself DID change.
    const active = document.activeElement;
    let focusedId = null, focusedValue = null, focusedSelStart = null, focusedSelEnd = null;
    if (active && active.id && active.id.startsWith('dc-')) {
      focusedId = active.id;
      focusedValue = active.value;
      focusedSelStart = active.selectionStart;
      focusedSelEnd = active.selectionEnd;
    }
    const pendingDcValues = {};
    document.querySelectorAll('[id^="dc-"]').forEach(el => {
      const id = el.id.slice(3);
      if (el.value !== '' && el.value !== '0') pendingDcValues[id] = el.value;
    });

    const main = document.getElementById('mainContent');
    if (orders.length === 0) {
      main.innerHTML = '<div class="empty-state"><div class="empty-icon">📭</div><div>No pending orders</div></div>';
      return;
    }
    main.innerHTML = orders.map(o => renderOrderCard(o, true)).join('');

    // Restore any in-progress delivery charge entries after the rebuild.
    Object.keys(pendingDcValues).forEach(id => {
      const el = document.getElementById('dc-' + id);
      if (el) el.value = pendingDcValues[id];
    });

    // Restore focus + cursor position so an in-progress keystroke isn't lost.
    if (focusedId) {
      const el = document.getElementById(focusedId);
      if (el) {
        el.value = focusedValue;
        el.focus();
        try { el.setSelectionRange(focusedSelStart, focusedSelEnd); } catch (e) {}
      }
    }
  }).catch(e => console.log('Load error', e));
}

function loadOrders() {
  if (currentTab === 'sales' || currentTab === 'settings') return;

  if (currentTab === 'pending') {
    checkForNewOrders(true); // true = force a render even if list is unchanged (e.g. on tab switch)
    return;
  }

  // history tab — browse a specific day's full order list, optionally
  // filtered by a search on order number / phone number. The date filter
  // (Today / date picker) and the search box now work together — e.g.
  // search a phone number within Today's orders to check for a missed
  // order — unless "All dates" is toggled on.
  const dateParam = historyDate || toDateStr(new Date());
  if (historySearchQuery) {
    const dateFilter = historySearchAllDates ? '' : '&date=' + dateParam;
    fetch('/api/search-orders?q=' + encodeURIComponent(historySearchQuery) + dateFilter).then(r => r.json()).then(orders => {
      renderHistoryList(orders, dateParam);
    }).catch(e => console.log('Search error', e));
    return;
  }
  fetch('/api/recent-orders?date=' + dateParam).then(r => r.json()).then(orders => {
    renderHistoryList(orders, dateParam);
  }).catch(e => console.log('Load error', e));
}

function renderHistoryList(orders, dateParam) {
  const main = document.getElementById('mainContent');
  const isToday = dateParam === toDateStr(new Date());
  const isSearching = !!historySearchQuery;

  const filtersHtml =
    '<div class="sales-filters">' +
      '<button class="sales-quick-btn ' + (isToday ? 'active' : '') + '" onclick="setHistoryDate(null)">Today</button>' +
      '<input type="date" id="historyDateInput" value="' + dateParam + '" onchange="setHistoryDate(this.value)">' +
    '</div>' +
    '<div class="history-search-row">' +
      '<input type="text" id="historySearchInput" class="history-search-input" placeholder="🔍 Search by order # or phone number..." value="' + historySearchQuery.replace(/"/g,'&quot;') + '" oninput="onHistorySearchInput(this.value)">' +
      (isSearching ? '<button class="history-search-clear" onclick="clearHistorySearch()">✕</button>' : '') +
    '</div>' +
    (isSearching ?
      '<label class="history-search-scope"><input type="checkbox" ' + (historySearchAllDates ? 'checked' : '') + ' onchange="toggleSearchScope(this.checked)"> Search all dates (currently: ' + (historySearchAllDates ? 'all dates' : (isToday ? 'today only' : dateParam + ' only')) + ')</label>'
      : '');

  if (orders.length === 0) {
    const scopeMsg = isSearching ? (historySearchAllDates ? '' : ' on ' + (isToday ? "today's orders" : dateParam)) : '';
    const emptyMsg = isSearching ? 'No orders found matching "' + historySearchQuery + '"' + scopeMsg : 'No orders on this date';
    main.innerHTML = filtersHtml + '<div class="empty-state"><div class="empty-icon">📋</div><div>' + emptyMsg + '</div></div>';
  } else {
    main.innerHTML = filtersHtml + orders.map(o => renderOrderCard(o, false)).join('');
  }

  // Re-focus the search box after rebuild so typing isn't interrupted,
  // same fix as applied to the pending-orders delivery-charge field.
  if (isSearching) {
    const input = document.getElementById('historySearchInput');
    if (input) {
      input.focus();
      const len = input.value.length;
      input.setSelectionRange(len, len);
    }
  }
}

let historySearchDebounce = null;
function onHistorySearchInput(value) {
  clearTimeout(historySearchDebounce);
  historySearchDebounce = setTimeout(() => {
    historySearchQuery = value.trim();
    loadOrders();
  }, 400);
}

function toggleSearchScope(checked) {
  historySearchAllDates = checked;
  loadOrders();
}

function clearHistorySearch() {
  historySearchQuery = '';
  historySearchAllDates = false;
  loadOrders();
}

function setHistoryDate(dateStr) {
  historyDate = dateStr || toDateStr(new Date());
  loadOrders();
}

function cancelOrder(id) {
  if (!confirm('Cancel this confirmed order? This cannot be undone.')) return;
  fetch('/api/cancel-order/' + id, {method:'POST'}).then(r => r.json()).then(res => {
    loadOrders();
  });
}

function confirmOrder(id) {
  const dcInput = document.getElementById('dc-' + id);
  const deliveryCharge = dcInput ? parseInt(dcInput.value) || 0 : 0;
  fetch('/api/confirm-order/' + id, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({delivery_charge: deliveryCharge})
  }).then(r => r.json()).then(res => {
    if (res.status === 'ok') {
      loadOrders();
    } else {
      alert('Print failed! Check printer connection.');
      loadOrders();
    }
  });
}

function rejectOrder(id) {
  if (!confirm('Reject this order?')) return;
  fetch('/api/reject-order/' + id, {method:'POST'}).then(r => r.json()).then(res => {
    loadOrders();
  });
}

function reprintOrder(id) {
  fetch('/api/reprint-order/' + id, {method:'POST'}).then(r => r.json()).then(res => {
    alert(res.status === 'ok' ? 'Reprinted!' : 'Print failed');
  });
}

function toDateStr(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function loadSales(quick) {
  const main = document.getElementById('mainContent');
  let start, end;

  if (quick === 'today') {
    start = end = toDateStr(new Date());
  } else if (quick === 'month') {
    const now = new Date();
    start = toDateStr(new Date(now.getFullYear(), now.getMonth(), 1));
    end = toDateStr(now);
  } else {
    // custom range: read whatever is currently in the date inputs, if present
    const sEl = document.getElementById('salesStart');
    const eEl = document.getElementById('salesEnd');
    start = (sEl && sEl.value) || toDateStr(new Date());
    end = (eEl && eEl.value) || toDateStr(new Date());
  }

  fetch('/api/sales-summary?start=' + start + '&end=' + end)
    .then(r => r.json())
    .then(summary => renderSales(summary, quick))
    .catch(e => { console.log('Sales load error', e); main.innerHTML = '<div class="sales-empty">Failed to load sales data.</div>'; });
}

function renderSales(summary, activeQuick) {
  const main = document.getElementById('mainContent');

  const filtersHtml =
    '<div class="sales-filters">' +
      '<button class="sales-quick-btn ' + (activeQuick === 'today' ? 'active' : '') + '" onclick="loadSales(\\'today\\')">Today</button>' +
      '<button class="sales-quick-btn ' + (activeQuick === 'month' ? 'active' : '') + '" onclick="loadSales(\\'month\\')">This Month</button>' +
      '<input type="date" id="salesStart" value="' + summary.start_date + '" onchange="loadSales(\\'range\\')">' +
      '<span style="color:var(--muted);font-size:0.8rem;">to</span>' +
      '<input type="date" id="salesEnd" value="' + summary.end_date + '" onchange="loadSales(\\'range\\')">' +
      '<a class="sales-quick-btn" style="text-decoration:none;" href="/api/export-orders-csv?start=' + summary.start_date + '&end=' + summary.end_date + '">⬇️ CSV</a>' +
    '</div>';

  if (summary.total_orders === 0) {
    main.innerHTML = filtersHtml + '<div class="sales-empty">📭 No confirmed orders in this range.</div>';
    return;
  }

  const byTypeHtml = summary.by_type.map(t =>
    '<div class="sales-row"><span>' + t.type + '</span><span>' + t.count + ' orders</span><span>₹' + t.revenue + '</span></div>'
  ).join('');

  const byDayHtml = summary.by_day.map(d =>
    '<div class="sales-row"><span>' + d.day + '</span><span>' + d.count + ' orders</span><span>₹' + d.revenue + '</span></div>'
  ).join('');

  const topItemsHtml = summary.top_items.map(i =>
    '<div class="sales-row"><span>' + i.name + '</span><span>x' + i.qty + '</span><span>₹' + i.revenue + '</span></div>'
  ).join('');

  main.innerHTML = filtersHtml +
    '<div class="sales-summary-card">' +
      '<div class="sales-total-label">Total Revenue</div>' +
      '<div class="sales-total-value">₹' + summary.total_revenue + '</div>' +
      '<div class="sales-total-sub">' + summary.total_orders + ' confirmed order' + (summary.total_orders === 1 ? '' : 's') + ' &middot; ' + summary.start_date + (summary.start_date !== summary.end_date ? ' to ' + summary.end_date : '') + '</div>' +
    '</div>' +
    '<div class="sales-section"><div class="sales-section-title">By Order Type</div>' + byTypeHtml + '</div>' +
    (summary.by_day.length > 1 ? '<div class="sales-section"><div class="sales-section-title">By Day</div>' + byDayHtml + '</div>' : '') +
    '<div class="sales-section"><div class="sales-section-title">Top Items</div>' + topItemsHtml + '</div>';
}

let currentSettings = { special_dates: {}, banner: { active: false, text: '', emoji: '🎉' } };
let dashPassword = '';

function loadSettings() {
  const main = document.getElementById('mainContent');
  fetch('/api/settings').then(r => r.json()).then(settings => {
    currentSettings = settings;
    renderSettings();
  }).catch(e => {
    console.log('Settings load error', e);
    main.innerHTML = '<div class="sales-empty">Failed to load settings.</div>';
  });
}

function renderSettings() {
  const main = document.getElementById('mainContent');
  const banner = currentSettings.banner || { active: false, text: '', emoji: '🎉' };
  const dates = currentSettings.special_dates || {};

  const bannerHtml =
    '<div class="settings-section">' +
      '<div class="settings-section-title">📣 Announcement Banner</div>' +
      '<div class="toggle-row"><span>Show banner on website</span><div class="toggle-switch ' + (banner.active ? 'on' : '') + '" id="bannerToggle" onclick="toggleBanner()"></div></div>' +
      '<label class="field-label">Emoji</label>' +
      '<input class="settings-input" id="bannerEmoji" maxlength="4" value="' + (banner.emoji || '🎉') + '" style="width:80px;">' +
      '<label class="field-label">Banner Text</label>' +
      '<input class="settings-input" id="bannerText" placeholder="e.g. Diwali Special: 20% off all pizzas!" value="' + (banner.text || '').replace(/"/g,'&quot;') + '">' +
    '</div>';

  const dateCards = Object.keys(dates).sort().map(dateKey => {
    const d = dates[dateKey];
    const isClosed = d.open === null;
    return '<div class="special-date-card" data-date="' + dateKey + '">' +
      '<div class="special-date-header">' +
        '<input type="date" class="settings-input" style="margin-bottom:0;width:auto;" value="' + dateKey + '" onchange="this.parentElement.parentElement.dataset.date=this.value">' +
        '<button class="btn-remove-date" onclick="removeSpecialDate(\\'' + dateKey + '\\')">✕ Remove</button>' +
      '</div>' +
      '<label class="field-label">Label (shown to customers)</label>' +
      '<input class="settings-input" placeholder="e.g. New Year\\'s Eve — Open till Midnight!" value="' + (d.label || '').replace(/"/g,'&quot;') + '" data-field="label">' +
      '<div class="toggle-row"><span>Force closed all day</span><div class="toggle-switch ' + (isClosed ? 'on' : '') + '" data-field="closed" onclick="this.classList.toggle(\\'on\\');this.parentElement.nextElementSibling.style.display=this.classList.contains(\\'on\\')?\\'none\\':\\'grid\\'"></div></div>' +
      '<div class="row-2col" style="display:' + (isClosed ? 'none' : 'grid') + ';">' +
        '<div><label class="field-label">Open (24h, e.g. 11 or 9.5)</label><input class="settings-input" type="number" step="0.5" value="' + (isClosed ? 11 : d.open) + '" data-field="open"></div>' +
        '<div><label class="field-label">Close (24h, e.g. 23.5, 24=midnight)</label><input class="settings-input" type="number" step="0.5" value="' + (isClosed ? 21.5 : d.close) + '" data-field="close"></div>' +
      '</div>' +
    '</div>';
  }).join('');

  const datesHtml =
    '<div class="settings-section">' +
      '<div class="settings-section-title">📅 Special Date Overrides (extended hours, closures)</div>' +
      dateCards +
      '<button class="btn-add-date" onclick="addSpecialDate()">+ Add Special Date</button>' +
    '</div>';

  const backupHtml =
    '<div class="settings-section">' +
      '<div class="settings-section-title">💾 Backups</div>' +
      '<div style="font-size:0.82rem;color:var(--muted);margin-bottom:0.7rem;line-height:1.5;">Your order database backs itself up automatically once a day on this PC. For real protection against this PC failing, download a copy now and then again periodically, and save it to Google Drive, a USB drive, or email it to yourself.</div>' +
      '<button class="btn-add-date" style="border-style:solid;color:var(--white);" onclick="downloadBackupNow()">⬇️ Download Backup Now</button>' +
    '</div>';

  main.innerHTML = bannerHtml + datesHtml + backupHtml +
    '<button class="btn-save" onclick="saveAllSettings()">💾 Save Settings</button>';
}

function downloadBackupNow() {
  if (!dashPassword) {
    dashPassword = prompt('Re-enter dashboard password to download a backup:') || '';
    if (!dashPassword) return;
  }
  window.location.href = '/api/download-backup?password=' + encodeURIComponent(dashPassword);
}

function toggleBanner() {
  document.getElementById('bannerToggle').classList.toggle('on');
}

function addSpecialDate() {
  const todayStr = toDateStr(new Date());
  if (!currentSettings.special_dates) currentSettings.special_dates = {};
  if (currentSettings.special_dates[todayStr]) { alert('That date already has an entry — scroll to edit it.'); return; }
  currentSettings.special_dates[todayStr] = { open: 11, close: 21.5, label: '' };
  renderSettings();
}

function removeSpecialDate(dateKey) {
  if (!confirm('Remove this special date override?')) return;
  delete currentSettings.special_dates[dateKey];
  renderSettings();
}

function collectSettingsFromForm() {
  const banner = {
    active: document.getElementById('bannerToggle').classList.contains('on'),
    text: document.getElementById('bannerText').value.trim(),
    emoji: document.getElementById('bannerEmoji').value.trim() || '🎉'
  };

  const special_dates = {};
  document.querySelectorAll('.special-date-card').forEach(card => {
    const dateKey = card.dataset.date;
    if (!dateKey) return;
    const label = card.querySelector('[data-field="label"]').value.trim();
    const closedToggle = card.querySelector('[data-field="closed"]');
    const isClosed = closedToggle.classList.contains('on');
    if (isClosed) {
      special_dates[dateKey] = { open: null, close: null, label: label };
    } else {
      const open = parseFloat(card.querySelector('[data-field="open"]').value);
      const close = parseFloat(card.querySelector('[data-field="close"]').value);
      special_dates[dateKey] = { open: open, close: close, label: label };
    }
  });

  return { banner, special_dates };
}

function saveAllSettings() {
  if (!dashPassword) {
    dashPassword = prompt('Re-enter dashboard password to save settings:') || '';
    if (!dashPassword) return;
  }
  const payload = collectSettingsFromForm();
  payload.password = dashPassword;

  fetch('/api/settings', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload)
  }).then(r => r.json().then(body => ({status: r.status, body}))).then(({status, body}) => {
    if (status === 200 && body.status === 'ok') {
      alert('Settings saved!');
      dashPassword = payload.password;
      loadSettings();
    } else {
      dashPassword = '';
      alert('Failed to save: ' + (body.message || 'Unknown error'));
    }
  }).catch(e => { alert('Failed to save settings.'); console.log(e); });
}

function startPolling() {
  loadOrders();
  checkForNewOrders(); // catch up immediately, don't wait for the first interval tick
  setInterval(checkForNewOrders, 5000); // always runs, regardless of active tab — this is the alarm
  setInterval(() => { if (currentTab !== 'pending') loadOrders(); }, 5000); // keeps history/search results fresh while viewing them
}

if (sessionStorage.getItem('gi_dash_auth') === '1') {
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('app').classList.add('visible');
  startPolling();
}
</script>
</body>
</html>
"""


if __name__ == "__main__":
    print("=" * 50)
    print("  Grill Inn — Order Server")
    print("  Dashboard: http://localhost:5000/dashboard")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000, debug=True)

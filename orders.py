from flask import Flask, request, Response
import win32print
import sqlite3
import json
from datetime import datetime

app = Flask(__name__)

# ── CONFIG ──────────────────────────────────────────────
PRINTER_NAME = "POS80 (1)"
DASHBOARD_PASSWORD = "grillinn2026"  # change this to whatever you like
DB_PATH = "orders.db"
# ────────────────────────────────────────────────────────

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


def save_order(data):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO orders (order_number, customer_name, customer_phone, order_type,
            address, notes, items_json, subtotal, packing_charge, delivery_charge,
            grand_total, order_time, status, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?)
    """, (
        data.get("order_number", ""),
        data.get("customer_name", ""),
        data.get("customer_phone", ""),
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


# ── DASHBOARD ROUTES ──────────────────────────────────────
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
  </div>
  <div class="main" id="mainContent"></div>
</div>

<script>
let currentTab = 'pending';
let knownPendingIds = new Set();
let isFirstLoad = true;
let historyDate = null;

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
  if (tab === 'sales') {
    loadSales('today');
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
      osc.type = 'sine';
      gain.gain.setValueAtTime(0.3, ctx.currentTime + start);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + start + dur);
      osc.start(ctx.currentTime + start);
      osc.stop(ctx.currentTime + start + dur);
    };
    playBeep(880, 0, 0.15);
    playBeep(1100, 0.18, 0.15);
    playBeep(880, 0.36, 0.25);
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
      ? '<div style="margin-bottom:0.5rem;"><label style="font-size:0.72rem;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:0.25rem;">Delivery Charge (₹)</label><input id="dc-' + order.id + '" type="number" min="0" value="0" style="width:100%;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--white);font-size:0.9rem;padding:0.5rem 0.75rem;outline:none;margin-bottom:0.5rem;" placeholder="Enter delivery charge..."></div>'
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

function loadOrders() {
  if (currentTab === 'sales') return;

  if (currentTab === 'pending') {
    fetch('/api/pending-orders').then(r => r.json()).then(orders => {
      const currentIds = new Set(orders.map(o => o.id));
      if (!isFirstLoad) {
        const newOnes = [...currentIds].filter(id => !knownPendingIds.has(id));
        if (newOnes.length > 0) playNotifySound();
      }
      knownPendingIds = currentIds;
      isFirstLoad = false;
      document.getElementById('pendingCount').textContent = orders.length;

      const main = document.getElementById('mainContent');
      if (orders.length === 0) {
        main.innerHTML = '<div class="empty-state"><div class="empty-icon">📭</div><div>No pending orders</div></div>';
        return;
      }
      main.innerHTML = orders.map(o => renderOrderCard(o, true)).join('');
    }).catch(e => console.log('Load error', e));
    return;
  }

  // history tab — browse a specific day's full order list
  const dateParam = historyDate || toDateStr(new Date());
  fetch('/api/recent-orders?date=' + dateParam).then(r => r.json()).then(orders => {
    renderHistoryList(orders, dateParam);
  }).catch(e => console.log('Load error', e));
}

function renderHistoryList(orders, dateParam) {
  const main = document.getElementById('mainContent');
  const isToday = dateParam === toDateStr(new Date());

  const filtersHtml =
    '<div class="sales-filters">' +
      '<button class="sales-quick-btn ' + (isToday ? 'active' : '') + '" onclick="setHistoryDate(null)">Today</button>' +
      '<input type="date" id="historyDateInput" value="' + dateParam + '" onchange="setHistoryDate(this.value)">' +
    '</div>';

  if (orders.length === 0) {
    main.innerHTML = filtersHtml + '<div class="empty-state"><div class="empty-icon">📋</div><div>No orders on this date</div></div>';
    return;
  }
  main.innerHTML = filtersHtml + orders.map(o => renderOrderCard(o, false)).join('');
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
      '<button class="sales-quick-btn ' + (activeQuick === 'today' ? 'active' : '') + '" onclick="loadSales(\'today\')">Today</button>' +
      '<button class="sales-quick-btn ' + (activeQuick === 'month' ? 'active' : '') + '" onclick="loadSales(\'month\')">This Month</button>' +
      '<input type="date" id="salesStart" value="' + summary.start_date + '" onchange="loadSales(\'range\')">' +
      '<span style="color:var(--muted);font-size:0.8rem;">to</span>' +
      '<input type="date" id="salesEnd" value="' + summary.end_date + '" onchange="loadSales(\'range\')">' +
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

function startPolling() {
  loadOrders();
  setInterval(loadOrders, 5000);
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

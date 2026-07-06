from flask import Flask, request, Response
import win32print
import sqlite3
import json
import re
import os
import shutil
import threading
import time
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from urllib.parse import quote
try:
    import qrcode
    from PIL import Image
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

app = Flask(__name__)

# ── CONFIG ──────────────────────────────────────────────
PRINTER_NAME = "POS80 (1)"
DASHBOARD_PASSWORD = "0000"  # change this to whatever you like
DB_PATH = "orders.db"
SETTINGS_PATH = "settings.json"

# Email receipts — sent only after you confirm an order on the dashboard,
# since that's when the delivery charge and final total are known.
# Uses Gmail SMTP, which is free. To set this up:
#   1. Use a Gmail account (create a dedicated one for the restaurant if you like).
#   2. Turn on 2-Step Verification on that Google account.
#   3. Create an "App Password" at https://myaccount.google.com/apppasswords
#      (NOT your normal Gmail password — a 16-character app-specific one).
#   4. Paste that below. Leave SMTP_EMAIL blank to disable email receipts entirely.
SMTP_EMAIL = "grillinnfalkawn@gmail.com"            # e.g. "grillinnfalkawn@gmail.com"
SMTP_APP_PASSWORD = "gervpwotqmzohgvb"     # 16-character Gmail App Password
RESTAURANT_NAME = "Grill Inn, Falkawn"
RESTAURANT_PHONE = "9612992023"

# UPI payment QR on receipts — requests the exact bill amount, with the
# order number as the payment note, so it's clear which order it's for.
UPI_PAYEE_VPA = "Q171786848@ybl"
UPI_PAYEE_NAME = "PhonePeMerchant"
UPI_MCC = "0000"

# Daily sales report — emails a summary of the day's sales automatically
# (uses the same Gmail SMTP settings as receipt emails above). Leave
# OWNER_REPORT_EMAIL blank to disable this entirely. Time is 24h "HH:MM".
# NOTE: this only fires if orders.py is actually running at that time —
# if the PC is off or asleep at 21:30, that day's report won't send.
OWNER_REPORT_EMAIL = ""    # e.g. "youremail@gmail.com" — where to send the daily report
DAILY_REPORT_TIME = "21:30"
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
    },
    "stock": {
        "out_of_stock_groups": [],  # e.g. ["pizza_pan", "burgers"]
        "out_of_stock_items": []    # individual item ids, e.g. ["BEV001", "DES004"]
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
            merged.setdefault("stock", DEFAULT_SETTINGS["stock"])
            merged["stock"].setdefault("out_of_stock_groups", [])
            merged["stock"].setdefault("out_of_stock_items", [])
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
DOUBLE_HEIGHT_ONLY = GS + b'!\x01'  # taller but not wider — a step down from full double-size, and keeps character width normal so right-aligned columns still line up
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
            customer_email TEXT,
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
            created_at TEXT,
            source TEXT DEFAULT 'website',
            discount_percent INTEGER DEFAULT 0,
            discount_amount INTEGER DEFAULT 0,
            void_reason TEXT,
            reject_reason TEXT,
            client_ref TEXT
        )
    """)
    # Migrations for databases created before these columns existed.
    c.execute("PRAGMA table_info(orders)")
    existing_cols = [row[1] for row in c.fetchall()]
    for col, coltype, default in [
        ("customer_email", "TEXT", None),
        ("source", "TEXT", "'website'"),
        ("discount_percent", "INTEGER", "0"),
        ("discount_amount", "INTEGER", "0"),
        ("void_reason", "TEXT", None),
        ("reject_reason", "TEXT", None),
        ("client_ref", "TEXT", None),
    ]:
        if col not in existing_cols:
            default_clause = f" DEFAULT {default}" if default is not None else ""
            c.execute(f"ALTER TABLE orders ADD COLUMN {col} {coltype}{default_clause}")
            print(f"[DB MIGRATION] Added {col} column")
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


# ── MENU DATA (shared with the customer-facing site) ─────────────────────
# Single source of truth for the POS: item names, prices, categories.
# Keep in sync with the MENU object in index.html if the menu changes.
MENU_JSON = '''{"Veg Burgers":[{"id":"BUR001","title":"Aloo Tikki Burger","description":"Veg Burger","price":52},{"id":"BUR002","title":"Veg Surprise Burger","description":"Veg Burger","price":94},{"id":"BUR003","title":"Chilli Lava Burger","description":"Veg Burger","price":105},{"id":"BUR004","title":"Crunchy Corn Burger","description":"Veg Burger","price":105},{"id":"BUR005","title":"Crispy Paneer Burger","description":"Veg Burger","price":125},{"id":"BUR006","title":"Paneer Tikka Burger","description":"Veg Burger","price":125},{"id":"BUR007","title":"Peri Peri Paneer Burger","description":"Veg Burger","price":125},{"id":"BUR008","title":"Maha Veggie Burger","description":"Veg Burger","price":135},{"id":"ADD003","title":"Extra Cheese (Burger)","description":"Add-on","price":27}],"Non Veg Burgers":[{"id":"BUR009","title":"Fried Chicken Burger","description":"Chicken Burger","price":105},{"id":"BUR010","title":"Chicken Surprise Burger","description":"Chicken Burger","price":115},{"id":"BUR011","title":"Chicken Chilli Lava Burger","description":"Chicken Burger","price":136},{"id":"BUR012","title":"Tandoori Chicken Burger","description":"Chicken Burger","price":146},{"id":"BUR013","title":"Peri Peri Chicken Burger","description":"Chicken Burger","price":146},{"id":"BUR014","title":"Premium Fried Chicken Burger","description":"Chicken Burger","price":146},{"id":"BUR015","title":"Maha Chicken Burger","description":"Chicken Burger","price":157},{"id":"ADD004","title":"Extra Cheese (Burger)","description":"Add-on","price":27}],"Wraps":[{"id":"WRP001","title":"Crispy Paneer Wrap","description":"Veg Wrap","price":136},{"id":"WRP002","title":"Tandoori Paneer Tikka Wrap","description":"Veg Wrap","price":136},{"id":"WRP003","title":"Crispy Chicken Wrap","description":"Chicken Wrap","price":146},{"id":"WRP004","title":"Tandoori Chicken Tikka Wrap","description":"Chicken Wrap","price":146}],"Fried Chicken":[{"id":"FCH001","title":"Fried Chicken (2 Pieces)","description":"Fried Chicken","price":188},{"id":"FCH002","title":"Fried Chicken (4 Pieces)","description":"Fried Chicken","price":356},{"id":"FCH003","title":"Fried Chicken (6 Pieces)","description":"Fried Chicken","price":535},{"id":"FCH004","title":"Fried Chicken (9 Pieces)","description":"Fried Chicken","price":745}],"Grilled Chicken":[{"id":"GRC001","title":"Tandoori Grilled Chicken - Half (4 Pieces)","description":"Grilled Chicken","price":349},{"id":"GRC002","title":"Tandoori Grilled Chicken - Full (8 Pieces)","description":"Grilled Chicken","price":649}],"Veg Footlongs":[{"id":"FTL001","title":"Veggie Delight Footlong","description":"Veg Footlong","price":125},{"id":"FTL002","title":"Paneer Tikka Footlong","description":"Veg Footlong","price":146},{"id":"FTL003","title":"Spicy Paneer Footlong","description":"Veg Footlong","price":146},{"id":"FTL004","title":"Deluxe Veggie Footlong","description":"Veg Footlong","price":157},{"id":"FTL005","title":"Peri Peri Paneer Footlong","description":"Veg Footlong","price":157},{"id":"FTL006","title":"Veg Extravaganza Footlong","description":"Veg Footlong","price":157},{"id":"FTL007","title":"Veg Cheese Burst Footlong","description":"Veg Footlong","price":167}],"Non Veg Footlongs":[{"id":"FTL008","title":"Simply Chicken Footlong","description":"Chicken Footlong","price":146},{"id":"FTL009","title":"Chicken Tikka Footlong","description":"Chicken Footlong","price":157},{"id":"FTL010","title":"Spicy Chicken Footlong","description":"Chicken Footlong","price":157},{"id":"FTL011","title":"Deluxe Chicken Footlong","description":"Chicken Footlong","price":167},{"id":"FTL012","title":"Peri Peri Chicken Footlong","description":"Chicken Footlong","price":167},{"id":"FTL013","title":"Chicken Extravaganza Footlong","description":"Chicken Footlong","price":167},{"id":"FTL014","title":"Chicken Cheese Burst Footlong","description":"Chicken Footlong","price":177}],"Veg Sandwiches":[{"id":"SAN001","title":"Veg Grilled Sandwich","description":"Veg Sandwich","price":105},{"id":"SAN002","title":"Tandoori Paneer Tikka Sandwich","description":"Veg Sandwich","price":125},{"id":"SAN003","title":"Italian Veg Sandwich","description":"Veg Sandwich","price":136}],"Non Veg Sandwiches":[{"id":"SAN004","title":"Chicken Grilled Sandwich","description":"Chicken Sandwich","price":136},{"id":"SAN005","title":"Tandoori Chicken Tikka Sandwich","description":"Chicken Sandwich","price":136},{"id":"SAN006","title":"Italian Chicken Sandwich","description":"Chicken Sandwich","price":146}],"Veg Pizza":[{"id":"PIZ001","title":"Margherita Pizza - Pan (6 Inch)","description":"Veg Pizza","price":95},{"id":"PIZ002","title":"Margherita Pizza - Regular (9 Inch)","description":"Veg Pizza","price":146},{"id":"PIZ003","title":"Veggie Delight Pizza - Pan (6 Inch)","description":"Veg Pizza","price":115},{"id":"PIZ004","title":"Veggie Delight Pizza - Regular (9 Inch)","description":"Veg Pizza","price":199},{"id":"PIZ005","title":"Tandoori Paneer Tikka Pizza - Pan (6 Inch)","description":"Veg Pizza","price":146},{"id":"PIZ006","title":"Tandoori Paneer Tikka Pizza - Regular (9 Inch)","description":"Veg Pizza","price":250},{"id":"PIZ007","title":"Teekha Paneer Pizza - Pan (6 Inch)","description":"Veg Pizza","price":146},{"id":"PIZ008","title":"Teekha Paneer Pizza - Regular (9 Inch)","description":"Veg Pizza","price":250},{"id":"PIZ009","title":"Indi Paneer Pizza - Pan (6 Inch)","description":"Veg Pizza","price":146},{"id":"PIZ010","title":"Indi Paneer Pizza - Regular (9 Inch)","description":"Veg Pizza","price":250},{"id":"PIZ011","title":"Peri Peri Paneer Pizza - Pan (6 Inch)","description":"Veg Pizza","price":146},{"id":"PIZ012","title":"Peri Peri Paneer Pizza - Regular (9 Inch)","description":"Veg Pizza","price":250},{"id":"PIZ013","title":"Deluxe Veggie Pizza - Pan (6 Inch)","description":"Veg Pizza","price":157},{"id":"PIZ014","title":"Deluxe Veggie Pizza - Regular (9 Inch)","description":"Veg Pizza","price":262},{"id":"PIZ015","title":"Mushroom Delight Pizza - Pan (6 Inch)","description":"Veg Pizza","price":157},{"id":"PIZ016","title":"Mushroom Delight Pizza - Regular (9 Inch)","description":"Veg Pizza","price":262},{"id":"PIZ017","title":"Veg Extravaganza Pizza - Pan (6 Inch)","description":"Veg Pizza","price":178},{"id":"PIZ018","title":"Veg Extravaganza Pizza - Regular (9 Inch)","description":"Veg Pizza","price":292},{"id":"PIZ019","title":"Veg Overloaded Pizza - Pan (6 Inch)","description":"Veg Pizza","price":199},{"id":"PIZ020","title":"Veg Overloaded Pizza - Regular (9 Inch)","description":"Veg Pizza","price":314},{"id":"ADDPIZ001","title":"Cheese Burst Add-on - Pan (6 Inch)","description":"Add-on for Pan Pizza only","price":89,"is_addon":true,"addon_size":"pan"},{"id":"ADDPIZ002","title":"Cheese Burst Add-on - Regular (9 Inch)","description":"Add-on for Regular Pizza only","price":109,"is_addon":true,"addon_size":"regular"}],"Non Veg Pizza":[{"id":"PIZ021","title":"Chicken Delight Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":136},{"id":"PIZ022","title":"Chicken Delight Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":229},{"id":"PIZ023","title":"Tandoori Chicken Tikka Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":157},{"id":"PIZ024","title":"Tandoori Chicken Tikka Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":272},{"id":"PIZ025","title":"Teekha Chicken Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":157},{"id":"PIZ026","title":"Teekha Chicken Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":272},{"id":"PIZ027","title":"Indi Chicken Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":157},{"id":"PIZ028","title":"Indi Chicken Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":272},{"id":"PIZ029","title":"Peri Peri Chicken Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":157},{"id":"PIZ030","title":"Peri Peri Chicken Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":272},{"id":"PIZ031","title":"Deluxe Chicken Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":167},{"id":"PIZ032","title":"Deluxe Chicken Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":282},{"id":"PIZ033","title":"Chicken Seekh Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":167},{"id":"PIZ034","title":"Chicken Seekh Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":282},{"id":"PIZ035","title":"Chicken Extravaganza Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":199},{"id":"PIZ036","title":"Chicken Extravaganza Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":335},{"id":"PIZ037","title":"Chicken Overloaded Pizza - Pan (6 Inch)","description":"Chicken Pizza","price":219},{"id":"PIZ038","title":"Chicken Overloaded Pizza - Regular (9 Inch)","description":"Chicken Pizza","price":345},{"id":"ADDPIZ003","title":"Cheese Burst Add-on - Pan (6 Inch)","description":"Add-on for Pan Pizza only","price":89,"is_addon":true,"addon_size":"pan"},{"id":"ADDPIZ004","title":"Cheese Burst Add-on - Regular (9 Inch)","description":"Add-on for Regular Pizza only","price":109,"is_addon":true,"addon_size":"regular"}],"Pastas":[{"id":"PAS001","title":"Veg Arrabbiata Penne","description":"Veg Pasta","price":136},{"id":"PAS002","title":"Veg Creamy White Penne Pasta","description":"Veg Pasta","price":146},{"id":"PAS003","title":"Mix Sauce Veg Penne Pasta","description":"Veg Pasta","price":146},{"id":"PAS004","title":"Chicken Arrabbiata Penne","description":"Chicken Pasta","price":157},{"id":"PAS005","title":"Chicken Creamy White Penne Pasta","description":"Chicken Pasta","price":167},{"id":"PAS006","title":"Mix Sauce Chicken Penne Pasta","description":"Chicken Pasta","price":167}],"Garlic Breads":[{"id":"GRB001","title":"Garlic Bread with Cheese","description":"Veg Garlic Bread","price":115},{"id":"GRB002","title":"Garlic Bread Supreme","description":"Veg Garlic Bread","price":125},{"id":"GRB003","title":"Chicken Garlic Bread with Cheese","description":"Chicken Garlic Bread","price":125},{"id":"GRB004","title":"Chicken Garlic Bread Supreme","description":"Chicken Garlic Bread","price":135}],"Veg Snacks":[{"id":"SNK001","title":"Paneer Pops (6 Pieces)","description":"Veg Snack","price":115},{"id":"SNK002","title":"Paneer Pops (16 Pieces)","description":"Veg Snack","price":209},{"id":"SNK003","title":"Paneer Strips (4 Pieces)","description":"Veg Snack","price":125},{"id":"SNK004","title":"Paneer Strips (8 Pieces)","description":"Veg Snack","price":219},{"id":"FRI001","title":"French Fries - Regular","description":"Fries","price":62},{"id":"FRI002","title":"French Fries - Large","description":"Fries","price":83},{"id":"FRI003","title":"Peri Peri Fries - Regular","description":"Fries","price":83},{"id":"FRI004","title":"Peri Peri Fries - Large","description":"Fries","price":104},{"id":"FRI005","title":"Hot & Spicy Fries - Regular","description":"Fries","price":83},{"id":"FRI006","title":"Hot & Spicy Fries - Large","description":"Fries","price":104},{"id":"FRI007","title":"Peri Peri Overloaded Fries","description":"Loaded Fries","price":125},{"id":"FRI008","title":"Hot & Spicy Overloaded Fries","description":"Loaded Fries","price":125},{"id":"FRI009","title":"Crispy Paneer Overloaded Fries","description":"Loaded Fries","price":146}],"Non Veg Snacks":[{"id":"SNK005","title":"Chicken Nuggets (6 Pieces)","description":"Chicken Snack","price":136},{"id":"SNK006","title":"Chicken Nuggets (12 Pieces)","description":"Chicken Snack","price":250},{"id":"SNK007","title":"Chicken Popcorn (8 Pieces)","description":"Chicken Snack","price":136},{"id":"SNK008","title":"Chicken Popcorn (16 Pieces)","description":"Chicken Snack","price":252},{"id":"SNK009","title":"Chicken Strips (3 Pieces)","description":"Chicken Snack","price":125},{"id":"FRI010","title":"Crispy Chicken Overloaded Fries","description":"Loaded Fries","price":157}],"Beverages":[{"id":"BEV001","title":"Coke","description":"Soft Drink","price":41},{"id":"BEV002","title":"Sprite","description":"Soft Drink","price":41},{"id":"BEV003","title":"Fanta","description":"Soft Drink","price":41},{"id":"BEV004","title":"Iced Tea Lemon","description":"Cold Beverage","price":75},{"id":"BEV005","title":"Fresh Lime Soda","description":"Cold Beverage","price":95},{"id":"BEV006","title":"Mojito","description":"Cold Beverage","price":95},{"id":"BEV007","title":"Mango Delight","description":"Beverage","price":95},{"id":"BEV008","title":"Strawberry Blast","description":"Beverage","price":95},{"id":"BEV009","title":"Cafe Frappe","description":"Beverage","price":105},{"id":"BEV010","title":"Mango Shake","description":"Milkshake","price":105},{"id":"BEV011","title":"Strawberry Shake","description":"Milkshake","price":105},{"id":"BEV012","title":"Chocolate Shake","description":"Milkshake","price":105},{"id":"BEV013","title":"Butterscotch Shake","description":"Milkshake","price":105},{"id":"BEV014","title":"Oreo Shake","description":"Milkshake","price":125},{"id":"BEV015","title":"Kit Kat Shake","description":"Milkshake","price":125},{"id":"BEV016","title":"Packaged Drinking Water","description":"Drinking Water","price":10}],"Desserts":[{"id":"DES004","title":"Mango Ice Cream","description":"Dessert","price":58},{"id":"DES005","title":"Vanilla Ice Cream","description":"Dessert","price":58},{"id":"DES006","title":"Strawberry Ice Cream","description":"Dessert","price":58},{"id":"DES007","title":"Chocolate Ice Cream","description":"Dessert","price":58},{"id":"DES002","title":"Hot Chocolate Fudge","description":"Dessert","price":105},{"id":"DES003","title":"Fruit Sundae","description":"Dessert","price":105}],"Add Ons":[{"id":"ADD002","title":"Cheese Dip","description":"Add-on","price":30},{"id":"ADD005","title":"Extra Cheese","description":"Add-on","price":27}],"Combos":[{"id":"FMB001","title":"Veg Fun Meal Box","description":"Veg Surprise Burger, Reg. Fries, Paneer Pops (4 pcs), Coke, Perk Chocolate","price":279},{"id":"FMB002","title":"Non Veg Fun Meal Box","description":"Fried Chicken Burger, Reg. Fries, Chicken Popcorn (4 pcs), Coke, Perk Chocolate","price":289},{"id":"COM001","title":"Add 2 Pieces Garlic Bread + Coke","description":"Meal Combo Add-on","price":105},{"id":"COM002","title":"Add Fries + Coke Combo","description":"Meal Combo Add-on","price":119},{"id":"COM003","title":"Add 1 Piece Fried Chicken + Coke","description":"Meal Combo Add-on","price":135},{"id":"COM004","title":"Add Chicken Popcorn + Coke","description":"Meal Combo Add-on","price":159},{"id":"COM005","title":"Add 2 Pieces Garlic Bread Supreme + Coke","description":"Meal Combo Add-on","price":115}]}'''
MENU_DATA = json.loads(MENU_JSON)
ITEMS_BY_ID = {}
CATEGORY_BY_ID = {}
for _cat, _items in MENU_DATA.items():
    for _it in _items:
        ITEMS_BY_ID[_it["id"]] = _it
        CATEGORY_BY_ID[_it["id"]] = _cat


# ── OUT-OF-STOCK GROUPS ──────────────────────────────────────────────
# Group membership is computed dynamically from the menu (not hardcoded
# item ids) so it stays correct automatically if the menu ever changes.
# Beverages and Desserts are deliberately NOT included as groups here —
# those are toggled individually instead (see /api/stock-status).
def _grp_pizza_pan(item_id, item, cat):
    return item_id.startswith("PIZ") and ("- Pan" in item["title"] or item.get("addon_size") == "pan")

def _grp_pizza_regular(item_id, item, cat):
    return item_id.startswith("PIZ") and ("- Regular" in item["title"] or item.get("addon_size") == "regular")

def _grp_burgers(item_id, item, cat):
    return cat in ("Veg Burgers", "Non Veg Burgers") and not item_id.startswith("ADD")

def _grp_sandwiches(item_id, item, cat):
    return cat in ("Veg Sandwiches", "Non Veg Sandwiches")

def _grp_wraps(item_id, item, cat):
    return cat == "Wraps"

def _grp_pastas(item_id, item, cat):
    return cat == "Pastas"

def _grp_garlic_bread(item_id, item, cat):
    return cat == "Garlic Breads"

def _grp_fries(item_id, item, cat):
    return item_id.startswith("FRI")

def _grp_footlong(item_id, item, cat):
    return cat in ("Veg Footlongs", "Non Veg Footlongs")

def _grp_paneer(item_id, item, cat):
    text = (item["title"] + " " + item.get("description", "")).lower()
    return "paneer" in text

def _grp_tandoori_chicken(item_id, item, cat):
    return "tandoori chicken" in item["title"].lower()

STOCK_GROUPS = {
    "pizza_pan": ("Pizza — Pan (6\")", _grp_pizza_pan),
    "pizza_regular": ("Pizza — Regular (9\")", _grp_pizza_regular),
    "burgers": ("Burgers", _grp_burgers),
    "sandwiches": ("Sandwiches", _grp_sandwiches),
    "wraps": ("Wraps", _grp_wraps),
    "pastas": ("Pastas", _grp_pastas),
    "garlic_bread": ("Garlic Bread", _grp_garlic_bread),
    "fries": ("Fries (incl. Overloaded)", _grp_fries),
    "footlong": ("Footlongs", _grp_footlong),
    "paneer": ("All Paneer Items", _grp_paneer),
    "tandoori_chicken": ("All Tandoori Chicken Items", _grp_tandoori_chicken),
}


def group_item_ids(group_key):
    """All menu item ids belonging to a given stock group."""
    _, matcher = STOCK_GROUPS[group_key]
    return [iid for iid, item in ITEMS_BY_ID.items() if matcher(iid, item, CATEGORY_BY_ID[iid])]


def compute_out_of_stock_item_ids(stock_settings):
    """Expands the group toggles + individually-toggled items into the
    full set of item ids currently out of stock."""
    ids = set(stock_settings.get("out_of_stock_items", []))
    for group_key in stock_settings.get("out_of_stock_groups", []):
        if group_key in STOCK_GROUPS:
            ids.update(group_item_ids(group_key))
    return sorted(ids)


# ── PACKAGING CHARGE LOGIC (Python port — used by the POS) ──────────────
# Mirrors the JS logic in index.html exactly. Kept here so POS orders get
# the identical packing charge as website orders for the same cart.
LARGE_PIZZA_IDS = {"PIZ002","PIZ004","PIZ006","PIZ008","PIZ010","PIZ012","PIZ014","PIZ016","PIZ018","PIZ020","PIZ022","PIZ024","PIZ026","PIZ028","PIZ030","PIZ032","PIZ034","PIZ036","PIZ038","PIZ040","PIZ042"}
SMALL_BOX_IDS = {
    "PIZ001","PIZ003","PIZ005","PIZ007","PIZ009","PIZ011","PIZ013","PIZ015","PIZ017","PIZ019","PIZ021","PIZ023","PIZ025","PIZ027","PIZ029","PIZ031","PIZ033","PIZ035","PIZ037",
    "SAN001","SAN002","SAN003","SAN004","SAN005","SAN006",
    "GRB001","GRB002","GRB003","GRB004",
    "FTL001","FTL002","FTL003","FTL004","FTL005","FTL006","FTL007","FTL008","FTL009","FTL010","FTL011","FTL012","FTL013","FTL014"
}
WRAP_IDS = {"WRP001","WRP002","WRP003","WRP004"}
PASTA_IDS = {"PAS001","PAS002","PAS003","PAS004","PAS005","PAS006","FRI007","FRI008","FRI009","FRI010"}


def item_packaging_cost(item_id):
    if item_id in LARGE_PIZZA_IDS: return 10
    if item_id in SMALL_BOX_IDS: return 7
    if item_id in WRAP_IDS: return 6
    if item_id in PASTA_IDS: return 10
    return 0


def packaging_slab(actual_cost):
    if actual_cost <= 15: return 0
    if actual_cost <= 30: return 10
    if actual_cost <= 45: return 15
    if actual_cost <= 60: return 20
    if actual_cost <= 80: return 25
    if actual_cost <= 100: return 30
    return 35


def calc_packaging_charge(items, order_type):
    """items: list of {"id": ..., "qty": ...}"""
    if order_type == "Dine In":
        return 0
    base = 0
    large_pizza_count = 0
    for it in items:
        iid = it.get("id", "")
        qty = it.get("qty", 0)
        base += item_packaging_cost(iid) * qty
        if iid in LARGE_PIZZA_IDS:
            large_pizza_count += qty
    if large_pizza_count > 0:
        bags = -(-large_pizza_count // 3)  # ceil division
    else:
        bags = 1 if base > 20 else 0
    actual_cost = base + (bags * 9)
    return packaging_slab(actual_cost)


def save_order(data):
    """Assigns the order number here on the server — never trusts a number
    the browser might send — so numbering is a single, authoritative,
    ever-increasing sequence instead of each customer's browser making up
    its own (which caused every new browser/device to start back at 001,
    and could even let two different customers collide on the same
    number). See next_website_order_number() below.

    Idempotency: if the browser retries a submission (e.g. after a dropped
    connection where the request actually succeeded but the response
    never made it back), we must not create a duplicate kitchen order —
    but we also can't key on order_number+phone anymore, since the number
    is now assigned here, not sent by the client. Instead the browser
    sends a `client_ref`: a short code it generates once per checkout
    attempt and reuses across retries. Same client_ref within 30 minutes
    = the same checkout attempt, so we return the already-assigned order
    instead of creating a second one.
    """
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    customer_phone = data.get("customer_phone", "")
    client_ref = data.get("client_ref", "")

    if client_ref:
        c.execute("""
            SELECT id, order_number FROM orders
            WHERE client_ref = ?
              AND created_at >= datetime('now', '-30 minutes')
            LIMIT 1
        """, (client_ref,))
        existing = c.fetchone()
        if existing:
            conn.close()
            return {"id": existing[0], "order_number": existing[1]}

    order_number = next_website_order_number()

    c.execute("""
        INSERT INTO orders (order_number, customer_name, customer_phone, customer_email, order_type,
            address, notes, items_json, subtotal, packing_charge, delivery_charge,
            grand_total, order_time, status, created_at, client_ref)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
    """, (
        order_number,
        data.get("customer_name", ""),
        customer_phone,
        data.get("customer_email", ""),
        data.get("order_type", ""),
        data.get("address", ""),
        data.get("notes", ""),
        json.dumps(data.get("items", [])),
        data.get("subtotal", 0),
        data.get("packing_charge", 0),
        data.get("delivery_charge", 0),
        data.get("grand_total", 0),
        data.get("order_time", ""),
        datetime.now().isoformat(),
        client_ref
    ))
    conn.commit()
    order_id = c.lastrowid
    conn.close()
    return {"id": order_id, "order_number": order_number}


def _next_seq_for_today(source_filter, today_iso):
    """Shared helper: finds the last order placed today (matching the given
    source filter) and returns the next sequence number for that day.
    Uses a trailing-digit match rather than a fixed split, so it keeps
    working seamlessly even on the day the numbering format changes (an
    old-style plain number like '014' or 'P014' from earlier the same day
    is still read correctly and the count continues from it)."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(f"""
        SELECT order_number FROM orders
        WHERE ({source_filter}) AND created_at LIKE ?
        ORDER BY id DESC LIMIT 1
    """, (today_iso + "%",))
    row = c.fetchone()
    conn.close()
    if row and row[0]:
        m = re.search(r"(\d+)$", row[0])
        if m:
            return int(m.group(1)) + 1
    return 1


def next_website_order_number():
    """Website orders get their own date-prefixed sequence that resets
    each day: DDMMYY-001, DDMMYY-002, ... This keeps numbers short and
    unambiguous (no ever-growing digit count over the life of the
    business, and no collision between 'today's #005' and 'last week's
    #005' since the date is baked into the number itself). Assigned here
    on the server so it's a single authoritative counter — not something
    each customer's browser invents independently."""
    today = datetime.now()
    n = _next_seq_for_today("source != 'pos' OR source IS NULL", today.strftime("%Y-%m-%d"))
    return f"{today.strftime('%d%m%y')}-{str(n).zfill(3)}"


def next_pos_order_number():
    """POS orders get their own date-prefixed sequence with a 'P' prefix
    (PDDMMYY-001, PDDMMYY-002, ...) so they're visually distinct from
    website orders on receipts and in reports, and can be filtered/
    analyzed separately. Resets daily, same reasoning as above."""
    today = datetime.now()
    n = _next_seq_for_today("source = 'pos'", today.strftime("%Y-%m-%d"))
    return f"P{today.strftime('%d%m%y')}-{str(n).zfill(3)}"


def save_pos_order(order_number, customer_name, customer_phone, order_type, address, notes,
                    items, subtotal, packing_charge, delivery_charge, discount_percent, discount_amount, grand_total):
    """POS orders skip the pending queue entirely — they're inserted with
    status already 'printed' since the cashier confirms and prints in one
    step, standing right at the counter."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO orders (order_number, customer_name, customer_phone, customer_email, order_type,
            address, notes, items_json, subtotal, packing_charge, delivery_charge,
            grand_total, order_time, status, created_at, source, discount_percent, discount_amount)
        VALUES (?, ?, ?, '', ?, ?, ?, ?, ?, ?, ?, ?, ?, 'printed', ?, 'pos', ?, ?)
    """, (
        order_number, customer_name, customer_phone, order_type, address, notes,
        json.dumps(items), subtotal, packing_charge, delivery_charge, grand_total,
        datetime.now().strftime("%d/%m/%Y, %H:%M"), datetime.now().isoformat(),
        discount_percent, discount_amount
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


def get_top_selling_items(days=30, limit=8):
    """Returns the best-selling menu items (by quantity sold) from
    confirmed orders in the last `days` days — used to auto-populate the
    POS Quick Add strip instead of a hand-maintained list."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("""
        SELECT items_json FROM orders
        WHERE status = 'printed'
          AND created_at >= datetime('now', ?)
    """, (f'-{int(days)} days',))
    rows = c.fetchall()
    conn.close()

    qty_by_title = {}
    for row in rows:
        try:
            items = json.loads(row["items_json"] or "[]")
        except Exception:
            continue
        for it in items:
            title = it.get("product_retailer_id", "")
            qty = it.get("quantity", 0) or 0
            qty_by_title[title] = qty_by_title.get(title, 0) + qty

    # Match each sold title back to a canonical menu item (titles may have
    # a note appended in brackets, e.g. "Pizza - Regular (9 Inch) [extra cheese]").
    matched = {}
    for title, qty in qty_by_title.items():
        for menu_id, item in ITEMS_BY_ID.items():
            if title.startswith(item["title"]):
                if menu_id not in matched or qty > matched[menu_id]["qty"]:
                    matched[menu_id] = {"id": menu_id, "title": item["title"], "price": item["price"], "qty": qty}
                break

    top = sorted(matched.values(), key=lambda x: x["qty"], reverse=True)[:limit]
    return top


def mark_order_status(order_id, status):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE orders SET status = ? WHERE id = ?", (status, order_id))
    conn.commit()
    conn.close()


def void_order(order_id, reason):
    """Voids an already-confirmed/printed order (e.g. customer changed
    their mind after ringing it up at the POS). Kept in the database with
    the reason for your records, but excluded from sales totals since
    get_sales_summary() only counts status='printed'."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE orders SET status = 'voided', void_reason = ? WHERE id = ?", (reason, order_id))
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
    """Search orders by order number, phone number, or reference code
    (the code shown to a customer if their order failed to send but may
    have gone through anyway), newest first. If date_str is given,
    results are restricted to that single date; otherwise all dates."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    like_query = "%" + query.strip() + "%"
    if date_str:
        c.execute("""
            SELECT * FROM orders
            WHERE (order_number LIKE ? OR customer_phone LIKE ? OR client_ref LIKE ?)
              AND date(created_at) = date(?)
            ORDER BY id DESC
            LIMIT ?
        """, (like_query, like_query, like_query, date_str, limit))
    else:
        c.execute("""
            SELECT * FROM orders
            WHERE order_number LIKE ? OR customer_phone LIKE ? OR client_ref LIKE ?
            ORDER BY id DESC
            LIMIT ?
        """, (like_query, like_query, like_query, limit))
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
    by_source = {}
    item_stats = {}

    for o in rows:
        gt = o.get("grand_total") or 0
        total_revenue += gt

        otype = o.get("order_type") or "Unknown"
        t = by_type.setdefault(otype, {"count": 0, "revenue": 0})
        t["count"] += 1
        t["revenue"] += gt

        src = o.get("source") or "website"
        s = by_source.setdefault(src, {"count": 0, "revenue": 0})
        s["count"] += 1
        s["revenue"] += gt

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
    by_source_list = [{"source": k, "count": v["count"], "revenue": v["revenue"]} for k, v in by_source.items()]
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
        "by_source": by_source_list,
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


def image_to_escpos(img):
    """Converts a PIL 1-bit image into ESC/POS GS v 0 raster bitmap bytes."""
    width, height = img.size
    # Width must be a multiple of 8 (1 bit per pixel, packed into bytes).
    pad_w = (width + 7) // 8 * 8
    if pad_w != width:
        padded = Image.new("1", (pad_w, height), 1)  # white background
        padded.paste(img, (0, 0))
        img = padded
        width = pad_w
    pixels = img.load()
    width_bytes = width // 8
    data = bytearray()
    for y in range(height):
        for xb in range(width_bytes):
            byte = 0
            for bit in range(8):
                x = xb * 8 + bit
                if pixels[x, y] == 0:  # black pixel
                    byte |= (0x80 >> bit)
            data.append(byte)
    header = GS + b'v0' + bytes([0]) + bytes([
        width_bytes & 0xFF, (width_bytes >> 8) & 0xFF,
        height & 0xFF, (height >> 8) & 0xFF
    ])
    return header + bytes(data)


def generate_upi_qr_bytes(amount, order_number, size_mm=50, dots_per_mm=8):
    """Builds a UPI payment QR requesting the exact bill amount, with the
    order number as the payment note, and returns it as ready-to-print
    ESC/POS raster bytes. Returns None if the qrcode/Pillow libraries
    aren't installed, so a missing dependency never breaks printing."""
    if not QR_AVAILABLE:
        print("[QR WARNING] qrcode/Pillow not installed — skipping payment QR. Run: pip install qrcode[pil] pillow --break-system-packages")
        return None
    try:
        note = quote(f"Order {order_number}")
        upi_url = (
            f"upi://pay?pa={UPI_PAYEE_VPA}&pn={quote(UPI_PAYEE_NAME)}&mc={UPI_MCC}"
            f"&mode=02&purpose=00&am={amount}&cu=INR&tn={note}"
        )
        qr = qrcode.QRCode(border=1, box_size=10)
        qr.add_data(upi_url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("1")
        size_px = size_mm * dots_per_mm
        img = img.resize((size_px, size_px), Image.NEAREST)  # NEAREST keeps QR edges crisp
        return image_to_escpos(img)
    except Exception as e:
        print(f"[QR ERROR] Could not generate payment QR: {e}")
        return None


def format_kot_datetime(order_time_str):
    """Reformats order_time (stored as 'DD/MM/YYYY, HH:MM') into the
    'Jul 02 2026 06:34 PM' style used on the reference KOT layout. Falls
    back to the original string if it doesn't parse cleanly."""
    try:
        dt = datetime.strptime(order_time_str.strip(), "%d/%m/%Y, %H:%M")
        return dt.strftime("%b %d %Y %I:%M %p")
    except Exception:
        return order_time_str


def shorten_pizza_name(name):
    """Shortens pizza names for the RECEIPT only: drops the '(9 Inch)'/
    '(6 Inch)' size hint (that's only useful to the customer browsing
    online) and drops the dash before it, e.g. 'Indi Chicken Pizza - Pan
    (6 Inch)' becomes 'Indi Chicken Pizza Pan'. The KOT uses the more
    detailed kot_item_name() below instead, matching the old POS output."""
    name = name.replace(" (9 Inch)", "").replace(" (6 Inch)", "")
    name = name.replace(" - Regular", " Regular").replace(" - Pan", " Pan")
    return name


# Explicit overrides for KOT item names that don't follow a clean pattern
# — matched exactly against the customer-facing menu title, based on the
# old POS's reference printouts you provided.
KOT_NAME_OVERRIDES = {
    "Peri Peri Overloaded Fries": "Peri Peri Cheesy Fries",
    "Hot & Spicy Overloaded Fries": "Hot & Spicy Cheesy Fries",
    "Tandoori Grilled Chicken - Half (4 Pieces)": "Tandoori Grilled Chicken (Half)",
    "Tandoori Grilled Chicken - Full (8 Pieces)": "Tandoori Grilled Chicken (Full)",
    "Hot Chocolate Fudge": "Chocolate Fudge Dessert",
    "Fruit Sundae": "Fruit Sundae Dessert",
    # Combo add-ons — the menu/receipt keep the customer-facing wording;
    # only the KOT uses this shorthand the kitchen is used to.
    "Add 2 Pieces Garlic Bread + Coke": "Pay Rs 105 More & Add Garlic Bread Cheese + Coke",
    "Add 2 Pieces Garlic Bread Supreme + Coke": "Pay Rs 115 More & Add Garlic Bread Supreme + Coke",
    "Add Fries + Coke Combo": "Pay Rs 119 More & Add Fries + Coke",
    "Add 1 Piece Fried Chicken + Coke": "Add 1 pcs Fried Chicken + Coke",
    "Add Chicken Popcorn + Coke": "Add Chicken Popcorn 8 pcs + Coke",
}


def kot_item_name(name):
    """Renames items for the KOT to match the old POS's exact output
    (per the reference printouts) — the kitchen staff are used to these
    names, so the KOT should read the same even though the customer-
    facing site and receipt use the fuller/clearer names. Only affects
    the KOT; the receipt uses shorten_pizza_name() above instead."""
    # Split off any note in brackets first, e.g. "Pizza - Pan (6 Inch) [extra cheese]"
    note = ""
    if " [" in name and name.endswith("]"):
        name, note = name[:name.index(" [")], name[name.index(" ["):]

    if name in KOT_NAME_OVERRIDES:
        return KOT_NAME_OVERRIDES[name] + note

    # Pizza sizing: "X Pizza - Regular (9 Inch)" -> "X Pizza (Regular)"
    #               "X Pizza - Pan (6 Inch)"     -> "X Pizza Pan"
    name = name.replace(" - Regular (9 Inch)", " (Regular)")
    name = name.replace(" - Pan (6 Inch)", " Pan")

    # Footlong: old POS appends a bare "6" (half-foot indicator)
    if "Footlong" in name and "Footlong 6" not in name:
        name = name.replace("Footlong", "Footlong 6")

    # Fries: "X Fries - Regular"/"- Large" -> "X Fries (Regular)"/"(Large)"
    m = re.match(r"^(.*) Fries - (Regular|Large)$", name)
    if m:
        name = f"{m.group(1)} Fries ({m.group(2)})"
    name = name.replace("Hot & Spicy Fries (Regular)", "Hot N Spicy Fries (Regular)")
    name = name.replace("Hot & Spicy Fries (Large)", "Hot N Spicy Fries (Large)")
    name = name.replace("French Fries (Regular)", "Classic French Fries (Regular)")
    name = name.replace("French Fries (Large)", "Classic French Fries (Large)")

    # "(N Pieces)" -> "(N pcs)", and "Strips" items drop the parens entirely
    name = name.replace("Pieces)", "pcs)")
    if "Strips" in name:
        name = re.sub(r"\((\d+) pcs\)", r"\1 pcs", name)

    # Ice creams print as "<Flavor> Dessert" on the old POS
    name = re.sub(r"^(\w+) Ice Cream$", r"\1 Dessert", name)

    return name + note


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
    discount_amount = order.get("discount_amount", 0) or 0
    discount_percent = order.get("discount_percent", 0) or 0

    buf = bytearray()

    def add(data: bytes):
        buf.extend(data)

    # ═══════════════════════════════
    # RECEIPT
    # ═══════════════════════════════
    add(INIT)
    add(ALIGN_CENTER)
    add(BOLD_ON)
    add(encode("Grill Inn"))
    add(BOLD_OFF)
    add(encode(f"Falkawn, {RESTAURANT_PHONE}"))
    add(encode("GSTIN: 15ARKPVB080N1ZX"))
    add(LF)
    add(BOLD_ON)
    add(DOUBLE_ON)
    add(encode(order_type))
    add(DOUBLE_OFF)
    add(BOLD_OFF)
    add(LF)

    add(ALIGN_LEFT)
    add(encode(f"{'Order #: ' + order_number:<24}{order_time:>24}"))
    add(encode(f"Customer: {customer_name}"))
    add(encode(f"Phone   : {customer_phone}"))
    if order_type == "Delivery" and address:
        add(encode(f"Address : {address}"))

    add(encode(divider()))
    add(encode(f"{'Item':<24}{'Qty':>6}{'Rate':>8}{'Amt':>8}"))
    add(encode(divider()))

    for item in items:
        name = shorten_pizza_name(item.get("product_retailer_id", "Unknown"))
        qty = item.get("quantity", 1)
        price = item.get("item_price", 0)
        amt = qty * price
        if len(name) <= 22:
            # Fits comfortably on one line with the qty/rate/amount columns.
            add(encode(f"{name:<24}{qty:>6}{price:>8}{amt:>8}"))
        else:
            # Name is genuinely too long — print it on its own line, then
            # the qty/rate/amount columns right-aligned on the next line.
            add(encode(name[:48]))
            add(encode(f"{'':<24}{qty:>6}{price:>8}{amt:>8}"))

    add(encode(divider()))
    add(encode(f"{'Sub Total':<32}{subtotal:>16}"))
    if packing_charge > 0:
        add(encode(f"{'Packing Charges':<32}{packing_charge:>16}"))
    if order_type == "Delivery":
        dc_str = str(delivery_charge) if delivery_charge else "______"
        add(encode(f"{'Delivery Charges':<32}{dc_str:>16}"))
    if discount_amount > 0:
        discount_label = f"Discount ({discount_percent}%)"
        add(encode(f"{discount_label:<32}{'-'+str(discount_amount):>16}"))
    add(encode(divider()))
    grand = subtotal + packing_charge + (delivery_charge if order_type == "Delivery" else 0) - discount_amount
    add(BOLD_ON)
    add(DOUBLE_HEIGHT_ONLY)
    add(encode(f"{'Bill Total':<32}{grand:>16}"))
    add(DOUBLE_OFF)
    add(BOLD_OFF)
    add(encode("Inclusive of 5% CGST & SGST"))

    if notes:
        add(encode(f"Notes: {notes[:40]}"))

    # UPI payment QR — requests the exact bill amount, order number as note.
    qr_bytes = generate_upi_qr_bytes(grand, order_number, size_mm=30)
    if qr_bytes:
        add(ALIGN_CENTER)
        add(qr_bytes)
        add(BOLD_ON)
        add(encode("Scan here to pay"))
        add(BOLD_OFF)

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
    add(DOUBLE_ON)
    add(encode(order_type))
    add(DOUBLE_OFF)
    add(BOLD_OFF)
    add(LF)

    add(ALIGN_LEFT)
    if customer_name and customer_name != "Walk-in":
        add(encode(f"Customer: {customer_name}"))
    if customer_phone:
        add(encode(f"Phone   : {customer_phone}"))
    if order_type == "Delivery" and address:
        add(encode(f"Address : {address}"))
    if notes:
        add(encode(f"Notes  : {notes[:40]}"))
    add(LF)

    add(encode(divider()))
    total_qty = sum(item.get("quantity", 1) for item in items)
    add(encode(f"{'Order # ' + order_number:<24}{str(len(items)) + ' items (' + str(total_qty) + ' Qty)':>24}"))
    add(encode(f"{format_kot_datetime(order_time):<24}{'Manager Falkawn':>24}"))
    add(encode(divider()))

    for item in items:
        name = kot_item_name(item.get("product_retailer_id", "Unknown"))[:35]
        qty = item.get("quantity", 1)
        add(encode(f"{name:<36}{qty:>12}"))
        add(LF)  # extra vertical spacing between items, matching the reference KOT layout

    add(encode(divider()))
    if order.get("source") != "pos":
        add(ALIGN_CENTER)
        add(BOLD_ON)
        add(encode("** ONLINE ORDER **"))
        add(BOLD_OFF)
        add(LF)

    printed_now = datetime.now()
    printed_str = printed_now.strftime("%b %d %Y %I:%M:%S.") + f"{printed_now.microsecond // 1000:03d}" + printed_now.strftime(" %p")
    add(ALIGN_CENTER)
    add(encode(f"Printed On: {printed_str}"))
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
        result = save_order(data)
        print(f"[ORDER SAVED] DB ID #{result['id']} — Order #{result['order_number']} — PENDING CONFIRMATION")

        response = Response(
            json.dumps({"status": "ok", "order_id": result["id"], "order_number": result["order_number"]}),
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
@app.route("/api/stock-status", methods=["GET", "OPTIONS"])
def api_get_stock_status():
    if request.method == "OPTIONS":
        response = Response("", status=200)
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "GET, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    settings = load_settings()
    out_of_stock_ids = compute_out_of_stock_item_ids(settings["stock"])
    response = Response(json.dumps({"out_of_stock_item_ids": out_of_stock_ids}), status=200, mimetype="application/json")
    response.headers["Access-Control-Allow-Origin"] = "*"
    return response


@app.route("/api/stock-groups")
def api_get_stock_groups():
    """Returns the available groups (with labels + current item counts)
    plus the Beverages/Desserts items, for the dashboard Settings UI."""
    settings = load_settings()
    stock = settings["stock"]
    groups = [
        {"key": key, "label": label, "count": len(group_item_ids(key)), "active": key in stock["out_of_stock_groups"]}
        for key, (label, _) in STOCK_GROUPS.items()
    ]
    individual_categories = {}
    for cat in ("Beverages", "Desserts"):
        individual_categories[cat] = [
            {"id": it["id"], "title": it["title"], "active": it["id"] in stock["out_of_stock_items"]}
            for it in MENU_DATA.get(cat, [])
        ]
    return Response(json.dumps({"groups": groups, "individual": individual_categories}), status=200, mimetype="application/json")


@app.route("/api/stock-status", methods=["POST"])
def api_set_stock_status():
    data = request.get_json() or {}
    if data.get("password") != DASHBOARD_PASSWORD:
        return Response('{"status":"error","message":"Incorrect password"}', status=401, mimetype="application/json")

    settings = load_settings()
    if "out_of_stock_groups" in data:
        settings["stock"]["out_of_stock_groups"] = [g for g in data["out_of_stock_groups"] if g in STOCK_GROUPS]
    if "out_of_stock_items" in data:
        settings["stock"]["out_of_stock_items"] = [i for i in data["out_of_stock_items"] if i in ITEMS_BY_ID]
    save_settings(settings)
    return Response(json.dumps({"status": "ok"}), status=200, mimetype="application/json")


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
@app.route("/api/send-daily-report", methods=["POST"])
def api_send_daily_report():
    data = request.get_json() or {}
    if data.get("password") != DASHBOARD_PASSWORD:
        return Response('{"status":"error","message":"Incorrect password"}', status=401, mimetype="application/json")
    if not OWNER_REPORT_EMAIL or not SMTP_EMAIL or not SMTP_APP_PASSWORD:
        return Response('{"status":"error","message":"Email/report not configured in orders.py"}', status=400, mimetype="application/json")
    ok = send_daily_sales_report()
    return Response(json.dumps({"status": "ok" if ok else "error"}), status=200, mimetype="application/json")


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
    return DASHBOARD_HTML.replace("__MENU_JSON_PLACEHOLDER__", MENU_JSON)


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


def send_receipt_email(order):
    """Sends a plain-text receipt email once an order is confirmed. Silently
    does nothing if SMTP isn't configured or the customer didn't provide an
    email — this must never block order confirmation/printing."""
    to_email = (order.get("customer_email") or "").strip()
    if not to_email or not SMTP_EMAIL or not SMTP_APP_PASSWORD:
        return

    try:
        items = json.loads(order.get("items_json") or "[]")
    except Exception:
        items = []

    lines = []
    for i in items:
        qty = i.get("quantity", 0)
        name = i.get("product_retailer_id", "")
        price = i.get("item_price", 0)
        lines.append(f"  {qty} x {name} - Rs.{qty * price}")
    items_text = "\n".join(lines) if lines else "  (no items listed)"

    subtotal = order.get("subtotal", 0) or 0
    packing = order.get("packing_charge", 0) or 0
    delivery = order.get("delivery_charge", 0) or 0
    grand_total = order.get("grand_total", 0) or 0

    charge_lines = f"Subtotal: Rs.{subtotal}\n"
    if packing > 0:
        charge_lines += f"Packing Charges: Rs.{packing}\n"
    if order.get("order_type") == "Delivery":
        charge_lines += f"Delivery Charges: Rs.{delivery}\n"
    charge_lines += f"Grand Total: Rs.{grand_total}"

    body = f"""Hi {order.get('customer_name', '')},

Thank you for your order from {RESTAURANT_NAME}! Your order has been confirmed.

Order #{order.get('order_number', '')}
Order Type: {order.get('order_type', '')}
{"Delivery Address: " + order.get("address", "") if order.get("order_type") == "Delivery" else ""}

Items:
{items_text}

{charge_lines}

Payment: Cash/UPI on {"Delivery" if order.get("order_type") == "Delivery" else "Pickup/Arrival"}

Questions? Call us at {RESTAURANT_PHONE}.

Thanks for ordering with us!
{RESTAURANT_NAME}
""".strip()

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = f"Order Confirmed - #{order.get('order_number', '')} - {RESTAURANT_NAME}"
        msg["From"] = SMTP_EMAIL
        msg["To"] = to_email

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=10) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_APP_PASSWORD)
            server.send_message(msg)
        print(f"[EMAIL] Receipt sent to {to_email} for order #{order.get('order_number')}")
    except Exception as e:
        # Never let an email failure affect order confirmation/printing.
        print(f"[EMAIL ERROR] Could not send receipt to {to_email}: {e}")


def send_daily_sales_report(date_str=None):
    """Emails a summary of the day's sales (defaults to today) to
    OWNER_REPORT_EMAIL — total revenue, breakdown by order type and by
    source (Website vs POS), and top-selling items. Lets you check daily
    sales from anywhere without needing to open the dashboard."""
    if not OWNER_REPORT_EMAIL or not SMTP_EMAIL or not SMTP_APP_PASSWORD:
        return False

    date_str = date_str or datetime.now().strftime("%Y-%m-%d")
    summary = get_sales_summary(date_str, date_str)

    if summary["total_orders"] == 0:
        body = f"No confirmed orders were recorded on {date_str}."
    else:
        lines = [
            f"Total Orders: {summary['total_orders']}",
            f"Total Revenue: Rs.{summary['total_revenue']}",
            "",
            "By Order Type:",
        ]
        for t in summary["by_type"]:
            lines.append(f"  {t['type']}: {t['count']} orders - Rs.{t['revenue']}")
        lines.append("")
        lines.append("By Source:")
        for s in summary["by_source"]:
            label = "Website" if s["source"] == "website" else "POS (In-Restaurant)"
            lines.append(f"  {label}: {s['count']} orders - Rs.{s['revenue']}")
        lines.append("")
        lines.append("Top Items:")
        for it in summary["top_items"][:10]:
            lines.append(f"  {it['qty']} x {it['name']} - Rs.{it['revenue']}")
        body = "\n".join(lines)

    body = f"""Daily Sales Report — {RESTAURANT_NAME}
Date: {date_str}

{body}

(Automated report — sent from your order management system.)
""".strip()

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = f"Daily Sales Report - {date_str} - {RESTAURANT_NAME}"
        msg["From"] = SMTP_EMAIL
        msg["To"] = OWNER_REPORT_EMAIL

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=10) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_APP_PASSWORD)
            server.send_message(msg)
        print(f"[DAILY REPORT] Sent for {date_str} to {OWNER_REPORT_EMAIL}")
        return True
    except Exception as e:
        print(f"[DAILY REPORT ERROR] Could not send report: {e}")
        return False


def daily_report_loop():
    """Checks once a minute whether it's time to send the daily report,
    and sends it at most once per day even if this loop keeps running
    past that minute."""
    last_sent_date = None
    while True:
        now = datetime.now()
        current_time = now.strftime("%H:%M")
        today_str = now.strftime("%Y-%m-%d")
        if current_time == DAILY_REPORT_TIME and last_sent_date != today_str:
            send_daily_sales_report(today_str)
            last_sent_date = today_str
        time.sleep(30)


threading.Thread(target=daily_report_loop, daemon=True).start()


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

    # Send the receipt email in the background so a slow/unavailable SMTP
    # connection never delays confirming or printing the order.
    order["grand_total"] = grand_total
    threading.Thread(target=send_receipt_email, args=(order,), daemon=True).start()

    return Response(
        json.dumps({"status": "ok" if success else "print_failed"}),
        status=200,
        mimetype="application/json"
    )


@app.route("/api/top-sellers")
def api_top_sellers():
    days = request.args.get("days", 30, type=int)
    limit = request.args.get("limit", 8, type=int)
    top = get_top_selling_items(days=days, limit=limit)
    return Response(json.dumps(top), status=200, mimetype="application/json")


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
    data = request.get_json() or {}
    reason = (data.get("reason") or "").strip()
    if not reason:
        return Response('{"status":"error","message":"A reason is required to reject an order"}', status=400, mimetype="application/json")
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE orders SET status = 'rejected', reject_reason = ? WHERE id = ?", (reason, order_id))
    conn.commit()
    conn.close()
    return Response('{"status":"ok"}', status=200, mimetype="application/json")


@app.route("/api/restore-order/<int:order_id>", methods=["POST"])
def api_restore_order(order_id):
    """Undoes a reject — puts the order back into the pending queue.
    Used by the 5-second Undo toast so a fat-fingered Reject isn't final."""
    mark_order_status(order_id, "pending")
    return Response('{"status":"ok"}', status=200, mimetype="application/json")


@app.route("/api/cancel-order/<int:order_id>", methods=["POST"])
def api_cancel_order(order_id):
    order = get_order_by_id(order_id)
    if not order:
        return Response('{"status":"not_found"}', status=404, mimetype="application/json")
    mark_order_status(order_id, "cancelled")
    return Response('{"status":"ok"}', status=200, mimetype="application/json")


@app.route("/api/pos-order", methods=["POST"])
def api_pos_order():
    """In-restaurant POS checkout — skips the pending queue entirely and
    prints immediately, since the cashier is standing at the counter
    confirming the order in person."""
    data = request.get_json() or {}

    order_type = data.get("order_type", "Dine In")
    customer_name = (data.get("customer_name") or "").strip() or "Walk-in"
    customer_phone = (data.get("customer_phone") or "").strip()
    address = (data.get("address") or "").strip()
    notes = (data.get("notes") or "").strip()
    raw_items = data.get("items", [])  # [{id, qty, note}]
    discount_percent = max(0, min(100, int(data.get("discount_percent", 0) or 0)))
    delivery_charge = int(data.get("delivery_charge", 0) or 0)

    if order_type == "Delivery" and not customer_phone:
        return Response(json.dumps({"status": "error", "message": "Phone number required for Delivery orders"}), status=400, mimetype="application/json")
    if not raw_items:
        return Response(json.dumps({"status": "error", "message": "Cart is empty"}), status=400, mimetype="application/json")

    # Price everything server-side from the canonical menu — never trust
    # prices sent from the browser.
    items_for_receipt = []
    packaging_items = []
    subtotal = 0
    for ri in raw_items:
        iid = ri.get("id", "")
        qty = int(ri.get("qty", 0) or 0)
        if qty <= 0:
            continue
        menu_item = ITEMS_BY_ID.get(iid)
        if not menu_item:
            continue
        note = (ri.get("note") or "").strip()
        title = menu_item["title"] + (f" [{note}]" if note else "")
        price = menu_item["price"]
        subtotal += price * qty
        items_for_receipt.append({"product_retailer_id": title, "quantity": qty, "item_price": price})
        packaging_items.append({"id": iid, "qty": qty})

    if not items_for_receipt:
        return Response(json.dumps({"status": "error", "message": "No valid items in cart"}), status=400, mimetype="application/json")

    packing_charge = calc_packaging_charge(packaging_items, order_type)
    discount_amount = round(subtotal * discount_percent / 100)
    grand_total = subtotal + packing_charge - discount_amount + (delivery_charge if order_type == "Delivery" else 0)

    order_number = next_pos_order_number()
    order_id = save_pos_order(
        order_number, customer_name, customer_phone, order_type, address, notes,
        items_for_receipt, subtotal, packing_charge, delivery_charge,
        discount_percent, discount_amount, grand_total
    )

    order = get_order_by_id(order_id)
    success = print_order(order)

    return Response(json.dumps({
        "status": "ok" if success else "print_failed",
        "order_number": order_number,
        "grand_total": grand_total
    }), status=200, mimetype="application/json")


@app.route("/api/void-order/<int:order_id>", methods=["POST"])
def api_void_order(order_id):
    order = get_order_by_id(order_id)
    if not order:
        return Response('{"status":"not_found"}', status=404, mimetype="application/json")
    data = request.get_json() or {}
    reason = (data.get("reason") or "").strip()
    void_order(order_id, reason)
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
  header{background:var(--charcoal);border-bottom:2px solid var(--flame);padding:0.9rem 1rem;display:flex;align-items:center;justify-content:space-between;gap:1rem;position:sticky;top:0;z-index:50;flex-wrap:wrap;}
  .tabs{display:flex;gap:0.5rem;flex-wrap:wrap;}
  .tab-btn{background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);cursor:pointer;font-weight:700;padding:0.55rem 1.1rem;font-size:0.85rem;white-space:nowrap;}
  .tab-btn.active{background:var(--flame);border-color:var(--flame);color:white;}
  .main{max-width:800px;margin:0 auto;padding:1rem;}
  .main.pos-active{max-width:none;padding:0;}

  /* ── POS ── */
  .pos-wrap{display:flex;height:calc(100vh - 68px);}
  .pos-menu-col{flex:1;overflow-y:auto;padding:1rem;border-right:1px solid var(--border);}
  .pos-cart-col{width:30%;min-width:320px;max-width:460px;display:flex;flex-direction:column;background:var(--charcoal);}
  .pos-search{width:100%;background:var(--panel);border:1px solid var(--border);border-radius:10px;color:var(--white);font-size:1.05rem;padding:0.9rem 2.6rem 0.9rem 1.1rem;outline:none;font-family:'Inter',sans-serif;}
  .pos-search:focus{border-color:var(--orange);}
  .pos-search-wrap{position:relative;margin-bottom:0.9rem;}
  .pos-search-clear{position:absolute;right:0.6rem;top:50%;transform:translateY(-65%);background:var(--border);border:none;border-radius:50%;color:var(--white);width:26px;height:26px;font-size:0.8rem;cursor:pointer;display:flex;align-items:center;justify-content:center;}
  .pos-cat-tabs{display:flex;gap:0.5rem;overflow-x:auto;padding-bottom:0.8rem;margin-bottom:0.8rem;border-bottom:1px solid var(--border);}
  .pos-cat-btn{flex-shrink:0;background:var(--panel);border:1px solid var(--border);border-radius:22px;color:var(--muted);font-size:0.95rem;font-weight:700;padding:0.6rem 1.2rem;cursor:pointer;white-space:nowrap;}
  .pos-cat-btn.active{background:var(--flame);border-color:var(--flame);color:white;}
  .pos-item-row{display:flex;align-items:center;justify-content:space-between;background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:1.1rem 1.3rem;margin-bottom:0.65rem;cursor:pointer;}
  .pos-item-row.pos-item-oos{opacity:0.45;cursor:not-allowed;}
  .pos-item-oos-tag{font-size:0.78rem;color:var(--red);font-weight:700;}
  .pos-item-row:active{background:var(--charcoal);}
  .pos-item-left{display:flex;align-items:center;gap:0.8rem;flex:1;min-width:0;}
  .pos-item-name{font-size:1.15rem;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .pos-item-price{font-size:1.1rem;color:var(--orange);font-weight:700;margin-left:0.8rem;flex-shrink:0;}
  .pos-veg-dot{width:13px;height:13px;border-radius:50%;flex-shrink:0;border:2px solid;}
  .pos-veg-dot.veg{background:#22c55e;border-color:#22c55e;}
  .pos-veg-dot.nonveg{background:#ef4444;border-color:#ef4444;}
  .pos-cart-header{padding:0.9rem 1rem 0.6rem;border-bottom:1px solid var(--border);}
  .pos-cart-items{flex:1;overflow-y:auto;padding:0.6rem 0.8rem;}
  .pos-cart-item{background:var(--panel);border-radius:8px;padding:0.5rem 0.6rem;margin-bottom:0.5rem;}
  .pos-cart-item-top{display:flex;justify-content:space-between;align-items:flex-start;gap:0.4rem;}
  .pos-cart-item-name{font-size:0.78rem;font-weight:600;line-height:1.3;}
  .pos-cart-item-price{font-size:0.75rem;color:var(--orange);font-weight:700;white-space:nowrap;}
  .pos-qty-row{display:flex;align-items:center;gap:0.4rem;margin-top:0.4rem;}
  .pos-qty-btn{width:24px;height:24px;border-radius:6px;background:var(--charcoal);border:1px solid var(--border);color:var(--white);font-size:0.9rem;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;}
  .pos-qty-input{width:40px;text-align:center;background:var(--charcoal);border:1px solid var(--border);border-radius:6px;color:var(--white);font-size:0.8rem;padding:0.2rem;}
  .pos-cart-footer{border-top:1px solid var(--border);padding:0.8rem 1rem;}
  .pos-summary-row{display:flex;justify-content:space-between;font-size:0.8rem;color:var(--muted);margin-bottom:0.3rem;}
  .pos-summary-row.total{color:var(--white);font-weight:700;font-size:0.95rem;margin-top:0.4rem;padding-top:0.4rem;border-top:1px solid var(--border);}
  .pos-field{width:100%;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--white);font-size:0.82rem;padding:0.5rem 0.65rem;outline:none;margin-bottom:0.5rem;font-family:'Inter',sans-serif;}
  .pos-type-tabs{display:flex;gap:0.4rem;margin-bottom:0.5rem;}
  .pos-type-btn{flex:1;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);font-size:0.78rem;font-weight:700;padding:0.5rem;cursor:pointer;}
  .pos-type-btn.active{background:var(--flame);border-color:var(--flame);color:white;}
  .pos-action-row{display:flex;gap:0.4rem;margin-top:0.5rem;}
  .pos-btn-secondary{flex:1;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.55rem;cursor:pointer;font-size:0.78rem;}
  .pos-btn-confirm{width:100%;background:linear-gradient(135deg,var(--green),#16a34a);border:none;border-radius:8px;color:white;font-weight:700;padding:0.75rem;cursor:pointer;font-size:0.9rem;margin-top:0.5rem;}
  .pos-btn-confirm:disabled{opacity:0.5;cursor:not-allowed;}
  .pos-held-bar{display:flex;gap:0.4rem;overflow-x:auto;padding:0.5rem 1rem;border-bottom:1px solid var(--border);background:rgba(249,115,22,0.08);}
  .pos-held-chip{flex-shrink:0;background:var(--panel);border:1px solid var(--orange);border-radius:20px;color:var(--orange);font-size:0.75rem;font-weight:700;padding:0.35rem 0.7rem;cursor:pointer;white-space:nowrap;}
  .pos-empty-cart{text-align:center;color:var(--muted);font-size:0.8rem;padding:2rem 1rem;}
  .conn-dot{width:9px;height:9px;border-radius:50%;background:var(--muted);flex-shrink:0;}
  .conn-dot.conn-ok{background:#22c55e;}
  .conn-dot.conn-down{background:var(--red);}
  .overdue-tag{font-size:0.68rem;color:white;background:var(--red);font-weight:700;margin-left:0.5rem;padding:0.15rem 0.5rem;border-radius:10px;}
  .order-card.overdue{border-color:var(--red);box-shadow:0 0 0 1px var(--red);}
  .undo-toast{position:fixed;bottom:1.2rem;left:50%;transform:translateX(-50%) translateY(100px);background:#1f1f1f;border:1px solid var(--border);border-radius:10px;padding:0.7rem 1rem;display:flex;align-items:center;gap:1rem;color:white;font-size:0.85rem;box-shadow:0 4px 20px rgba(0,0,0,0.4);z-index:200;transition:transform 0.25s;}
  .undo-toast.show{transform:translateX(-50%) translateY(0);}
  .undo-toast button{background:var(--orange);border:none;color:white;font-weight:700;padding:0.4rem 0.9rem;border-radius:6px;cursor:pointer;font-size:0.8rem;}
  .pos-fav-label{font-size:0.72rem;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:0.4rem;}
  .pos-fav-strip{display:flex;gap:0.5rem;overflow-x:auto;padding-bottom:0.8rem;margin-bottom:0.4rem;}
  .pos-fav-btn{flex-shrink:0;background:var(--panel);border:1px solid var(--orange);border-radius:10px;padding:0.55rem 0.8rem;cursor:pointer;text-align:left;min-width:110px;}
  .pos-fav-name{font-size:0.82rem;font-weight:700;color:var(--white);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:130px;}
  .pos-fav-price{font-size:0.78rem;color:var(--orange);font-weight:700;margin-top:0.15rem;}
  .pos-preset-row{display:flex;gap:0.4rem;margin-bottom:0.5rem;}
  .pos-preset-btn{flex:1;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.45rem;cursor:pointer;font-size:0.8rem;}
  .pos-preset-btn.active{background:var(--flame);border-color:var(--flame);color:white;}
  .pos-details-toggle{width:100%;background:var(--panel);border:1px solid var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.6rem;cursor:pointer;font-size:0.8rem;text-align:left;margin:0.6rem 0 0.5rem;}
  .pos-details-panel{background:rgba(255,255,255,0.02);border:1px dashed var(--border);border-radius:8px;padding:0.6rem;margin-bottom:0.5rem;}
  .pos-qty-picker-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.6);z-index:300;align-items:center;justify-content:center;}
  .pos-qty-picker-box{background:var(--charcoal);border:1px solid var(--border);border-radius:14px;padding:1.2rem;width:90%;max-width:340px;}
  .pos-qty-picker-title{font-size:0.95rem;font-weight:700;color:var(--white);margin-bottom:0.9rem;text-align:center;}
  .pos-qty-picker-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:0.6rem;margin-bottom:0.9rem;}
  .pos-qty-picker-btn{background:var(--panel);border:1px solid var(--border);border-radius:10px;color:var(--white);font-size:1.2rem;font-weight:700;padding:0.8rem 0;cursor:pointer;}
  .pos-qty-picker-btn:active{background:var(--flame);}
  .pos-qty-picker-cancel{width:100%;background:none;border:1px solid var(--border);border-radius:8px;color:var(--muted);font-weight:700;padding:0.6rem;cursor:pointer;}
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
  .status-tag.voided{background:rgba(148,163,184,0.15);color:var(--muted);}
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
  .sales-row{display:grid;grid-template-columns:1fr 56px 76px;align-items:center;gap:0.5rem;padding:0.4rem 0;border-bottom:1px solid var(--border);font-size:0.88rem;}
  .sales-row:last-child{border-bottom:none;}
  .sales-row span:first-child{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .sales-row span:nth-child(2){text-align:right;color:var(--muted);}
  .sales-row span:nth-child(3){text-align:right;font-weight:600;}
  .sales-empty{text-align:center;color:var(--muted);padding:2rem 1rem;font-size:0.85rem;}
  .settings-section{background:var(--charcoal);border:1px solid var(--border);border-radius:12px;padding:1rem;margin-bottom:1rem;}
  .settings-section-title{font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:var(--muted);margin-bottom:0.8rem;}
  .stock-toggle-row{display:flex;align-items:center;gap:0.6rem;padding:0.5rem 0;font-size:0.85rem;color:var(--white);cursor:pointer;border-bottom:1px solid var(--border);}
  .stock-toggle-row input{width:18px;height:18px;cursor:pointer;flex-shrink:0;}
  .stock-count{color:var(--muted);font-size:0.75rem;}
  .stock-cat-label{font-size:0.72rem;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:var(--orange);margin:0.8rem 0 0.3rem;}
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
    <div class="logo" style="font-size:1.6rem;display:flex;align-items:center;gap:0.5rem;">Grill Inn <span id="connDot" class="conn-dot" title="Checking connection..."></span></div>
    <div class="tabs">
      <button class="tab-btn active" id="pendingTab" onclick="switchTab('pending')">Pending <span class="badge-count" id="pendingCount">0</span></button>
      <button class="tab-btn" id="historyTab" onclick="switchTab('history')">History</button>
      <button class="tab-btn" id="salesTab" onclick="switchTab('sales')">📊 Sales</button>
      <button class="tab-btn" id="posTab" onclick="switchTab('pos')">🧾 POS</button>
      <button class="tab-btn" id="settingsTab" onclick="switchTab('settings')">⚙️ Settings</button>
    </div>
  </header>
  <div class="main" id="mainContent"></div>
</div>

<script>
let currentTab = 'pending';
let knownPendingIds = new Set();
let isFirstLoad = true;
let historyDate = null;
let historySearchQuery = '';
let historySearchAllDates = false;

// ── POS ──────────────────────────────────────────────────────────────
const MENU = __MENU_JSON_PLACEHOLDER__;
const MENU_CATEGORIES = Object.keys(MENU);
let posCart = {};              // { itemId: {id, title, price, qty, note, isVeg} }
let posSearch = '';
let posOrderType = 'Dine In';
let posDiscountPercent = 0;
let posDeliveryCharge = 0;
let posDetailsExpanded = false; // customer detail fields collapsed by default so cart items stay visible without scrolling
let posLastRemovedItem = null; // for Undo
let posHeldOrders = [];        // [{label, cart, orderType, name, phone, address, notes, discountPercent}]

function posIsVeg(item) {
  const d = (item.description || '').toLowerCase();
  const t = (item.title || '').toLowerCase();
  return !(d.includes('chicken') || t.includes('chicken'));
}

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
  document.getElementById('posTab').classList.toggle('active', tab === 'pos');
  document.getElementById('settingsTab').classList.toggle('active', tab === 'settings');
  document.getElementById('mainContent').classList.toggle('pos-active', tab === 'pos');
  if (tab === 'sales') {
    loadSales('today');
  } else if (tab === 'settings') {
    loadSettings();
  } else if (tab === 'pos') {
    renderPOS();
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

  const discountAmount = order.discount_amount || 0;
  const grandTotal = order.subtotal + (order.packing_charge || 0) - discountAmount + (order.order_type === 'Delivery' ? (order.delivery_charge || 0) : 0);
  const typeClass = order.order_type.replace(' ', '-');
  const sourceTag = order.source === 'pos' ? '<span style="font-size:0.68rem;color:var(--orange);font-weight:700;margin-left:0.4rem;">🧾 POS</span>' : '';

  const minutesWaiting = isPending ? Math.floor((Date.now() - new Date(order.created_at).getTime()) / 60000) : 0;
  const isOverdue = isPending && minutesWaiting >= 3;
  const overdueTag = isOverdue ? '<span class="overdue-tag">⏰ Waiting ' + minutesWaiting + 'm</span>' : '';

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
    const statusLabel = order.status === 'printed' ? '✅ Printed' : order.status === 'rejected' ? '✕ Rejected' : order.status === 'voided' ? '🚫 Voided' : order.status === 'cancelled' ? '🚫 Cancelled' : '⚠️ Print Failed';
    actionsHtml = '<div style="display:flex;justify-content:space-between;align-items:center;gap:0.5rem;flex-wrap:wrap;">' +
      '<span class="status-tag ' + order.status + '">' + statusLabel + '</span>' +
      '<div style="display:flex;gap:0.5rem;">' +
      (order.status === 'printed' ? '<button class="btn-reject" onclick="cancelOrder(' + order.id + ')">🚫 Void</button>' : '') +
      '<button class="btn-reprint" onclick="reprintOrder(' + order.id + ')">🖨️ Reprint</button>' +
      '</div></div>' +
      (order.status === 'voided' && order.void_reason ? '<div style="font-size:0.75rem;color:var(--muted);margin-top:0.4rem;">Reason: ' + order.void_reason + '</div>' : '') +
      (order.status === 'rejected' && order.reject_reason ? '<div style="font-size:0.75rem;color:var(--muted);margin-top:0.4rem;">Reason: ' + order.reject_reason + '</div>' : '');
  }

  return '<div class="order-card ' + (isPending ? 'pending' : '') + (isOverdue ? ' overdue' : '') + '">' +
    '<div class="order-header"><span class="order-num">#' + order.order_number + sourceTag + overdueTag + '</span><span class="order-type-tag ' + typeClass + '">' + order.order_type + '</span></div>' +
    '<div class="order-meta">' + order.customer_name + ' • ' + order.customer_phone + (order.address ? '<br>📍 ' + order.address : '') + '<br>🕒 ' + order.order_time + '</div>' +
    '<div class="order-items">' + itemsHtml +
      (discountAmount > 0 ? '<div class="order-item-row" style="color:var(--orange);"><span>Discount (' + (order.discount_percent||0) + '%)</span><span>−₹' + discountAmount + '</span></div>' : '') +
      '<div class="order-total">Total: ₹' + grandTotal + '</div></div>' +
    (order.notes ? '<div style="font-size:0.8rem;color:var(--muted);margin-bottom:0.5rem;">📝 ' + order.notes + '</div>' : '') +
    actionsHtml +
    '</div>';
}

// Detects new pending orders and sounds the alarm — runs on its own
// interval regardless of which dashboard tab is currently open (this is
// what used to be missing: previously the check only ran while viewing
// the Pending tab, so switching to History silenced new-order alerts).
// If the Pending tab happens to be open, it also renders the list.
let lastOverdueAlertTime = 0;

function checkForNewOrders(forceRender) {
  fetch('/api/pending-orders').then(r => r.json()).then(orders => {
    setConnDot(true);
    const currentIds = new Set(orders.map(o => o.id));

    if (!isFirstLoad) {
      const newOnes = [...currentIds].filter(id => !knownPendingIds.has(id));
      if (newOnes.length > 0) playNotifySound();
    }

    // Nudge again if any order has been waiting 3+ minutes unconfirmed —
    // easy to miss during a rush when new-order alerts blend together.
    const now = Date.now();
    const overdue = orders.some(o => now - new Date(o.created_at).getTime() > 3 * 60 * 1000);
    if (overdue && now - lastOverdueAlertTime > 60 * 1000) {
      playNotifySound();
      lastOverdueAlertTime = now;
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
  }).catch(e => { console.log('Load error', e); setConnDot(false); });
}

function setConnDot(ok) {
  const dot = document.getElementById('connDot');
  if (!dot) return;
  dot.className = 'conn-dot ' + (ok ? 'conn-ok' : 'conn-down');
  dot.title = ok ? 'Connected' : 'Cannot reach the server — check the connection';
}

function loadOrders() {
  if (currentTab === 'sales' || currentTab === 'settings' || currentTab === 'pos') return;

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
  if (!confirm('Void this confirmed order? This cannot be undone.')) return;
  const reason = prompt('Reason for voiding this order (optional):') || '';
  fetch('/api/void-order/' + id, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ reason })
  }).then(r => r.json()).then(res => {
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
  let reason = prompt('Reason for rejecting this order (required):');
  if (reason === null) return; // cancelled
  reason = reason.trim();
  if (!reason) { alert('A reason is required to reject an order — this keeps rejections deliberate, not accidental.'); return; }

  fetch('/api/reject-order/' + id, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ reason })
  }).then(r => r.json()).then(res => {
    if (res.status !== 'ok') { alert(res.message || 'Could not reject order.'); return; }
    loadOrders();
    showUndoToast('Order rejected.', () => {
      fetch('/api/restore-order/' + id, {method:'POST'}).then(r => r.json()).then(() => loadOrders());
    });
  });
}

let undoToastTimer = null;
function showUndoToast(message, onUndo) {
  let toast = document.getElementById('undoToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'undoToast';
    toast.className = 'undo-toast';
    document.body.appendChild(toast);
  }
  toast.innerHTML = '<span>' + message + '</span><button onclick="undoToastAction()">Undo</button>';
  toast.classList.add('show');
  clearTimeout(undoToastTimer);
  window._undoToastCallback = onUndo;
  undoToastTimer = setTimeout(() => { toast.classList.remove('show'); }, 5000);
}
function undoToastAction() {
  if (window._undoToastCallback) window._undoToastCallback();
  const toast = document.getElementById('undoToast');
  if (toast) toast.classList.remove('show');
  clearTimeout(undoToastTimer);
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

  const reportHtml =
    '<div class="settings-section">' +
      '<div class="settings-section-title">📧 Daily Sales Report</div>' +
      '<div style="font-size:0.82rem;color:var(--muted);margin-bottom:0.7rem;line-height:1.5;">A daily sales summary auto-sends by email at 21:30 (set OWNER_REPORT_EMAIL in orders.py to enable). This only works while this PC is on and orders.py is running at that time.</div>' +
      '<button class="btn-add-date" style="border-style:solid;color:var(--white);" onclick="sendReportNow()">📧 Send Today\\'s Report Now</button>' +
    '</div>';

  main.innerHTML = bannerHtml + datesHtml + '<div id="stockSection"></div>' + backupHtml + reportHtml +
    '<button class="btn-save" onclick="saveAllSettings()">💾 Save Settings</button>';
  loadStockSection();
}

let stockGroupsCache = null;

function loadStockSection() {
  fetch('/api/stock-groups').then(r => r.json()).then(data => {
    stockGroupsCache = data;
    renderStockSection();
  }).catch(e => console.log('Stock groups load error', e));
}

function renderStockSection() {
  const el = document.getElementById('stockSection');
  if (!el || !stockGroupsCache) return;

  const groupsHtml = stockGroupsCache.groups.map(g =>
    '<label class="stock-toggle-row">' +
      '<input type="checkbox" ' + (g.active ? 'checked' : '') + ' onchange="toggleStockGroup(&quot;' + g.key + '&quot;, this.checked)">' +
      '<span>' + g.label + ' <span class="stock-count">(' + g.count + ' items)</span></span>' +
    '</label>'
  ).join('');

  const individualHtml = Object.entries(stockGroupsCache.individual).map(([cat, items]) =>
    '<div class="stock-cat-label">' + cat + '</div>' +
    items.map(it =>
      '<label class="stock-toggle-row">' +
        '<input type="checkbox" ' + (it.active ? 'checked' : '') + ' onchange="toggleStockItem(&quot;' + it.id + '&quot;, this.checked)">' +
        '<span>' + it.title + '</span>' +
      '</label>'
    ).join('')
  ).join('');

  el.innerHTML =
    '<div class="settings-section">' +
      '<div class="settings-section-title">📦 Out of Stock</div>' +
      '<div style="font-size:0.82rem;color:var(--muted);margin-bottom:0.7rem;line-height:1.5;">Toggling any of these takes effect immediately on the website and POS — no need to hit Save Settings below. Group toggles cover whole categories (e.g. all Pan-size pizzas); Beverages and Desserts are toggled individually.</div>' +
      groupsHtml +
      '<div style="margin-top:0.8rem;">' + individualHtml + '</div>' +
    '</div>';
}

function toggleStockGroup(key, checked) {
  if (!dashPassword) {
    dashPassword = prompt('Re-enter dashboard password:') || '';
    if (!dashPassword) { loadStockSection(); return; }
  }
  stockGroupsCache.groups.forEach(g => { if (g.key === key) g.active = checked; });
  const groups = stockGroupsCache.groups.filter(g => g.active).map(g => g.key);
  fetch('/api/stock-status', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ password: dashPassword, out_of_stock_groups: groups })
  }).then(r => r.json()).then(res => {
    if (res.status !== 'ok') { dashPassword = ''; alert(res.message || 'Failed to update.'); }
    loadStockSection();
  });
}

function toggleStockItem(id, checked) {
  if (!dashPassword) {
    dashPassword = prompt('Re-enter dashboard password:') || '';
    if (!dashPassword) { loadStockSection(); return; }
  }
  const allItems = Object.values(stockGroupsCache.individual).flat();
  allItems.forEach(it => { if (it.id === id) it.active = checked; });
  const items = allItems.filter(it => it.active).map(it => it.id);
  fetch('/api/stock-status', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ password: dashPassword, out_of_stock_items: items })
  }).then(r => r.json()).then(res => {
    if (res.status !== 'ok') { dashPassword = ''; alert(res.message || 'Failed to update.'); }
    loadStockSection();
  });
}

function sendReportNow() {
  if (!dashPassword) {
    dashPassword = prompt('Re-enter dashboard password to send the report:') || '';
    if (!dashPassword) return;
  }
  fetch('/api/send-daily-report', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ password: dashPassword })
  }).then(r => r.json().then(body => ({status: r.status, body}))).then(({status, body}) => {
    if (status === 200 && body.status === 'ok') {
      alert('Report sent!');
    } else {
      dashPassword = '';
      alert('Failed to send: ' + (body.message || 'Check that OWNER_REPORT_EMAIL and SMTP settings are configured in orders.py'));
    }
  }).catch(e => alert('Failed to send report.'));
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

// ── POS: rendering ──────────────────────────────────────────────────
let posOutOfStockIds = new Set();

function fetchPosStockStatus() {
  fetch('/api/stock-status').then(r => r.json()).then(data => {
    posOutOfStockIds = new Set(data.out_of_stock_item_ids || []);
    renderPosItemList();
    renderPosFavorites();
  }).catch(e => console.log('Stock status load error', e));
}

function renderPOS() {
  const main = document.getElementById('mainContent');
  main.innerHTML =
    (posHeldOrders.length > 0 ?
      '<div class="pos-held-bar">' + posHeldOrders.map((h, i) =>
        '<div class="pos-held-chip" onclick="resumeHeldOrder(' + i + ')">↩ ' + h.label + '</div>'
      ).join('') + '</div>' : '') +
    '<div class="pos-wrap">' +
      '<div class="pos-menu-col">' +
        '<div class="pos-search-wrap">' +
          '<input type="text" class="pos-search" id="posSearchInput" placeholder="🔍 Search menu (any part of the name)..." value="' + esc2(posSearch) + '" oninput="onPosSearch(this.value)">' +
          (posSearch ? '<button class="pos-search-clear" onclick="clearPosSearch()">✕</button>' : '') +
        '</div>' +
        '<div id="posFavorites"></div>' +
        '<div id="posItemList"></div>' +
      '</div>' +
      '<div class="pos-cart-col">' +
        '<div class="pos-cart-header"><strong>🧾 Current Order</strong></div>' +
        '<div id="posOrderTypeBar"></div>' +
        '<div class="pos-cart-items" id="posCartItems"></div>' +
        '<div class="pos-cart-footer" id="posCartFooter"></div>' +
      '</div>' +
    '</div>' +
    '<div id="posQtyPickerOverlay" class="pos-qty-picker-overlay" style="display:none;"></div>';

  renderPosFavorites();
  renderPosItemList();
  renderPosOrderTypeBar();
  renderPosCart();
  fetchPosStockStatus();
}

// Top sellers are fetched from actual sales (last 7 days) — see
// renderPosFavorites() below — rather than a hand-maintained list, so
// this automatically tracks what's actually moving at the counter.
let posFavoritesCache = null;

function renderPosFavorites() {
  const el = document.getElementById('posFavorites');
  if (!el) return;

  if (posFavoritesCache === null) {
    el.innerHTML = '<div class="pos-fav-label">⭐ Quick Add (loading...)</div>';
    fetch('/api/top-sellers?days=30&limit=8').then(r => r.json()).then(items => {
      posFavoritesCache = items;
      renderPosFavorites();
    }).catch(e => { posFavoritesCache = []; renderPosFavorites(); });
    return;
  }

  const availableFavs = posFavoritesCache.filter(it => !posOutOfStockIds.has(it.id));
  if (availableFavs.length === 0) { el.innerHTML = ''; return; }
  el.innerHTML = '<div class="pos-fav-label">⭐ Quick Add — top sellers, last 30 days</div><div class="pos-fav-strip">' +
    availableFavs.map(it =>
      '<button class="pos-fav-btn" data-id="' + it.id + '" onclick="posItemClick(this.dataset.id)" onmousedown="posStartLongPress(this.dataset.id)" ontouchstart="posStartLongPress(this.dataset.id)" onmouseup="posEndLongPress()" ontouchend="posEndLongPress()">' +
        '<div class="pos-fav-name">' + it.title + '</div><div class="pos-fav-price">₹' + it.price + '</div>' +
      '</button>'
    ).join('') + '</div>';
}

function esc2(str) { return (str || '').replace(/"/g, '&quot;'); }

function onPosSearch(value) {
  posSearch = value;
  renderPosItemList();
  refreshPosSearchClearBtn();
}

function clearPosSearch() {
  posSearch = '';
  document.getElementById('posSearchInput').value = '';
  document.getElementById('posSearchInput').focus();
  renderPosItemList();
  refreshPosSearchClearBtn();
}

function refreshPosSearchClearBtn() {
  const wrap = document.querySelector('.pos-search-wrap');
  if (!wrap) return;
  let btn = wrap.querySelector('.pos-search-clear');
  if (posSearch && !btn) {
    btn = document.createElement('button');
    btn.className = 'pos-search-clear';
    btn.textContent = '✕';
    btn.onclick = clearPosSearch;
    wrap.appendChild(btn);
  } else if (!posSearch && btn) {
    btn.remove();
  }
}

function renderPosItemList() {
  const el = document.getElementById('posItemList');
  const q = posSearch.trim().toLowerCase();
  let items = [];
  MENU_CATEGORIES.forEach(cat => (MENU[cat] || []).forEach(it => items.push(it)));
  if (q) items = items.filter(it => it.title.toLowerCase().includes(q));

  if (items.length === 0) {
    el.innerHTML = '<div class="pos-empty-cart">No items found.</div>';
    return;
  }
  el.innerHTML = items.map(it => {
    const veg = posIsVeg(it);
    const oos = posOutOfStockIds.has(it.id);
    const handlers = oos ? '' : 'onclick="posItemClick(this.dataset.id)" onmousedown="posStartLongPress(this.dataset.id)" ontouchstart="posStartLongPress(this.dataset.id)" onmouseup="posEndLongPress()" ontouchend="posEndLongPress()"';
    return '<div class="pos-item-row' + (oos ? ' pos-item-oos' : '') + '" data-id="' + it.id + '" ' + handlers + '>' +
      '<div class="pos-item-left"><div class="pos-veg-dot ' + (veg ? 'veg' : 'nonveg') + '"></div>' +
      '<div class="pos-item-name">' + it.title + '</div></div>' +
      (oos ? '<div class="pos-item-oos-tag">Out of Stock</div>' : '<div class="pos-item-price">₹' + it.price + '</div>') +
    '</div>';
  }).join('');
}

// ── POS: cart management ────────────────────────────────────────────
let posLongPressTimer = null;
let posLongPressFired = false;

function posStartLongPress(id) {
  posLongPressFired = false;
  clearTimeout(posLongPressTimer);
  posLongPressTimer = setTimeout(() => {
    posLongPressFired = true;
    showPosQtyPicker(id);
  }, 450);
}

function posEndLongPress() {
  clearTimeout(posLongPressTimer);
}

function posItemClick(id) {
  // A long-press already opened the quantity picker for this tap — don't
  // also add a single qty on release.
  if (posLongPressFired) { posLongPressFired = false; return; }
  posAddItem(id);
}

function findMenuItem(id) {
  for (const cat of MENU_CATEGORIES) {
    const found = MENU[cat].find(i => i.id === id);
    if (found) return found;
  }
  return null;
}

function showPosQtyPicker(id) {
  const item = findMenuItem(id);
  if (!item) return;
  const overlay = document.getElementById('posQtyPickerOverlay');
  if (!overlay) return;
  const options = [1, 2, 3, 4, 5, 6, 8, 10];
  overlay.innerHTML =
    '<div class="pos-qty-picker-box">' +
      '<div class="pos-qty-picker-title">' + item.title + '</div>' +
      '<div class="pos-qty-picker-grid">' +
        options.map(n => '<button class="pos-qty-picker-btn" onclick="posQtyPickerChoose(&quot;' + id + '&quot;,' + n + ')">' + n + '</button>').join('') +
      '</div>' +
      '<button class="pos-qty-picker-cancel" onclick="hidePosQtyPicker()">Cancel</button>' +
    '</div>';
  overlay.style.display = 'flex';
  overlay.onclick = (e) => { if (e.target === overlay) hidePosQtyPicker(); };
}

function hidePosQtyPicker() {
  const overlay = document.getElementById('posQtyPickerOverlay');
  if (overlay) overlay.style.display = 'none';
}

function posQtyPickerChoose(id, qty) {
  const item = findMenuItem(id);
  if (!item) return;
  posCart[id] = { id: item.id, title: item.title, price: item.price, qty: qty, note: (posCart[id] && posCart[id].note) || '' };
  hidePosQtyPicker();
  clearPosSearchAfterAdd();
  renderPosCart();
}

function posRepeatLastOrder() {
  captureCartFields();
  const phone = (posCustomerPhone || '').trim();
  if (!phone) { alert('Enter a phone number first to find their last order.'); return; }
  fetch('/api/search-orders?q=' + encodeURIComponent(phone)).then(r => r.json()).then(orders => {
    const match = orders.find(o => o.customer_phone === phone);
    if (!match) { alert('No previous order found for that phone number.'); return; }
    if (Object.keys(posCart).length > 0 && !confirm('Replace the current cart with their last order (' + match.order_number + ')?')) return;
    let items;
    try { items = JSON.parse(match.items_json); } catch(e) { items = []; }
    const newCart = {};
    items.forEach(i => {
      const found = findMenuItemByTitle(i.product_retailer_id);
      if (found) newCart[found.id] = { id: found.id, title: found.title, price: found.price, qty: i.quantity, note: '' };
    });
    if (Object.keys(newCart).length === 0) { alert('Could not match any items from that order to the current menu.'); return; }
    posCart = newCart;
    posOrderType = match.order_type || 'Dine In';
    posCustomerName = match.customer_name || '';
    posAddress = match.address || '';
    posDetailsExpanded = true; // keep it open so the cashier can review what got filled in
    renderPOS();
  }).catch(e => alert('Could not fetch last order — check the connection.'));
}

function findMenuItemByTitle(title) {
  for (const cat of MENU_CATEGORIES) {
    const found = MENU[cat].find(i => title.startsWith(i.title));
    if (found) return found;
  }
  return null;
}

function posAddItem(id) {
  if (posOutOfStockIds.has(id)) { alert('This item is currently marked Out of Stock.'); return; }
  let item = null;
  for (const cat of MENU_CATEGORIES) {
    const found = MENU[cat].find(i => i.id === id);
    if (found) { item = found; break; }
  }
  if (!item) return;
  if (posCart[id]) posCart[id].qty += 1;
  else posCart[id] = { id: item.id, title: item.title, price: item.price, qty: 1, note: '' };
  clearPosSearchAfterAdd();
  renderPosCart();
}

// Clears the search box once an item's been added, so the cashier doesn't
// have to manually clear it before searching for the next item.
function clearPosSearchAfterAdd() {
  if (!posSearch) return;
  posSearch = '';
  const input = document.getElementById('posSearchInput');
  if (input) input.value = '';
  renderPosItemList();
  refreshPosSearchClearBtn();
}

function posChangeQty(id, delta) {
  const it = posCart[id];
  if (!it) return;
  it.qty += delta;
  if (it.qty <= 0) posRemoveItem(id, false);
  else renderPosCart();
}

function posSetQty(id, value) {
  const qty = parseInt(value) || 0;
  const it = posCart[id];
  if (!it) return;
  if (qty <= 0) posRemoveItem(id, false);
  else { it.qty = qty; renderPosCart(); }
}

function posRemoveItem(id, trackForUndo) {
  if (trackForUndo !== false && posCart[id]) posLastRemovedItem = { ...posCart[id] };
  delete posCart[id];
  renderPosCart();
}

function posUndoLastItem() {
  if (!posLastRemovedItem) { alert('Nothing to undo.'); return; }
  posCart[posLastRemovedItem.id] = posLastRemovedItem;
  posLastRemovedItem = null;
  renderPosCart();
}

function posClearCart() {
  if (Object.keys(posCart).length === 0) return;
  if (!confirm('Clear the current order?')) return;
  posCart = {};
  posDiscountPercent = 0;
  posDetailsExpanded = false;
  renderPosCart();
}

function posCalcTotals() {
  const items = Object.values(posCart);
  const subtotal = items.reduce((s, i) => s + i.price * i.qty, 0);
  const packagingItems = items.map(i => ({ id: i.id, qty: i.qty }));
  const packingCharge = calcPosPackaging(packagingItems, posOrderType);
  const discountAmount = Math.round(subtotal * posDiscountPercent / 100);
  const grandTotal = subtotal + packingCharge - discountAmount + (posOrderType === 'Delivery' ? posDeliveryCharge : 0);
  return { subtotal, packingCharge, discountAmount, grandTotal };
}

// Same packaging logic as the customer site / backend, kept in sync manually.
const POS_LARGE_PIZZA_IDS = new Set(['PIZ002','PIZ004','PIZ006','PIZ008','PIZ010','PIZ012','PIZ014','PIZ016','PIZ018','PIZ020','PIZ022','PIZ024','PIZ026','PIZ028','PIZ030','PIZ032','PIZ034','PIZ036','PIZ038','PIZ040','PIZ042']);
const POS_SMALL_BOX_IDS = new Set(['PIZ001','PIZ003','PIZ005','PIZ007','PIZ009','PIZ011','PIZ013','PIZ015','PIZ017','PIZ019','PIZ021','PIZ023','PIZ025','PIZ027','PIZ029','PIZ031','PIZ033','PIZ035','PIZ037','SAN001','SAN002','SAN003','SAN004','SAN005','SAN006','GRB001','GRB002','GRB003','GRB004','FTL001','FTL002','FTL003','FTL004','FTL005','FTL006','FTL007','FTL008','FTL009','FTL010','FTL011','FTL012','FTL013','FTL014']);
const POS_WRAP_IDS = new Set(['WRP001','WRP002','WRP003','WRP004']);
const POS_PASTA_IDS = new Set(['PAS001','PAS002','PAS003','PAS004','PAS005','PAS006','FRI007','FRI008','FRI009','FRI010']);
function posItemPackagingCost(id) {
  if (POS_LARGE_PIZZA_IDS.has(id)) return 10;
  if (POS_SMALL_BOX_IDS.has(id)) return 7;
  if (POS_WRAP_IDS.has(id)) return 6;
  if (POS_PASTA_IDS.has(id)) return 10;
  return 0;
}
function posPackagingSlab(actual) {
  if (actual <= 15) return 0;
  if (actual <= 30) return 10;
  if (actual <= 45) return 15;
  if (actual <= 60) return 20;
  if (actual <= 80) return 25;
  if (actual <= 100) return 30;
  return 35;
}
function calcPosPackaging(items, orderType) {
  if (orderType === 'Dine In') return 0;
  let base = 0, largeCount = 0;
  items.forEach(i => {
    base += posItemPackagingCost(i.id) * i.qty;
    if (POS_LARGE_PIZZA_IDS.has(i.id)) largeCount += i.qty;
  });
  const bags = largeCount > 0 ? Math.ceil(largeCount / 3) : (base > 20 ? 1 : 0);
  return posPackagingSlab(base + bags * 9);
}

function renderPosOrderTypeBar() {
  const el = document.getElementById('posOrderTypeBar');
  if (!el) return;
  el.innerHTML =
    '<div class="pos-type-tabs">' +
      ['Dine In', 'Takeaway', 'Delivery'].map(t =>
        '<button class="pos-type-btn ' + (posOrderType === t ? 'active' : '') + '" onclick="setPosOrderType(&quot;' + t + '&quot;)">' + t + '</button>'
      ).join('') +
    '</div>';
}

function renderPosCart() {
  const itemsEl = document.getElementById('posCartItems');
  const footerEl = document.getElementById('posCartFooter');
  if (!itemsEl || !footerEl) return; // not on POS tab

  const items = Object.values(posCart);
  if (items.length === 0) {
    itemsEl.innerHTML = '<div class="pos-empty-cart">Cart is empty.<br>Tap items on the left to add them.</div>';
  } else {
    itemsEl.innerHTML = items.map(it =>
      '<div class="pos-cart-item">' +
        '<div class="pos-cart-item-top"><div class="pos-cart-item-name">' + it.title + '</div><div class="pos-cart-item-price">₹' + (it.price * it.qty) + '</div></div>' +
        '<div class="pos-qty-row">' +
          '<button class="pos-qty-btn" onclick="posChangeQty(&quot;' + it.id + '&quot;,-1)">−</button>' +
          '<input class="pos-qty-input" type="number" value="' + it.qty + '" onchange="posSetQty(&quot;' + it.id + '&quot;,this.value)">' +
          '<button class="pos-qty-btn" onclick="posChangeQty(&quot;' + it.id + '&quot;,1)">+</button>' +
          '<button class="pos-qty-btn" style="margin-left:auto;color:var(--red);" onclick="posRemoveItem(&quot;' + it.id + '&quot;)">✕</button>' +
        '</div>' +
      '</div>'
    ).join('');
  }

  const { subtotal, packingCharge, discountAmount, grandTotal } = posCalcTotals();
  const deliveryPresets = [30, 40, 50];

  const detailsPreview = [posCustomerName, posCustomerPhone].filter(Boolean).join(', ') || 'optional';
  const detailsFields =
    '<input class="pos-field" id="posName" placeholder="Name' + (posOrderType === 'Delivery' ? ' *' : ' (optional)') + '" value="' + esc2(posCustomerName || '') + '">' +
    '<div style="display:flex;gap:0.5rem;">' +
      '<input class="pos-field" id="posPhone" style="flex:1;" placeholder="Phone' + (posOrderType === 'Delivery' ? ' *' : ' (optional)') + '" value="' + esc2(posCustomerPhone || '') + '">' +
      '<button class="pos-btn-secondary" style="flex:0 0 auto;padding:0 0.7rem;" onclick="posRepeatLastOrder()" title="Load this phone number\\'s last order">🔁</button>' +
    '</div>' +
    (posOrderType === 'Delivery' ? '<input class="pos-field" id="posAddress" placeholder="Delivery Address *" value="' + esc2(posAddress || '') + '">' : '') +
    (posOrderType === 'Delivery' ? '<input class="pos-field" id="posDeliveryCharge" type="number" placeholder="Delivery Charge (₹)" value="' + (posDeliveryCharge || '') + '" onchange="posDeliveryCharge=parseInt(this.value)||0;renderPosCart();">' : '') +
    (posOrderType === 'Delivery' ? '<div class="pos-preset-row">' + deliveryPresets.map(p =>
        '<button class="pos-preset-btn ' + (posDeliveryCharge === p ? 'active' : '') + '" onclick="posSetDeliveryCharge(' + p + ')">₹' + p + '</button>'
      ).join('') + '<button class="pos-preset-btn ' + (deliveryPresets.indexOf(posDeliveryCharge) === -1 ? 'active' : '') + '" onclick="posFocusDeliveryCharge()">Custom</button></div>' : '') +
    '<input class="pos-field" id="posNotes" placeholder="Notes (optional)" value="' + esc2(posNotes || '') + '">' +
    '<div class="pos-field" style="display:flex;align-items:center;gap:0.5rem;padding:0.4rem 0.65rem;">' +
      '<span style="font-size:0.78rem;color:var(--muted);white-space:nowrap;">Discount %</span>' +
      '<input type="number" min="0" max="100" value="' + posDiscountPercent + '" style="width:100%;background:transparent;border:none;color:var(--white);outline:none;text-align:right;" onchange="posDiscountPercent=Math.max(0,Math.min(100,parseInt(this.value)||0));renderPosCart();">' +
    '</div>';

  footerEl.innerHTML =
    '<div class="pos-summary-row"><span>Subtotal</span><span>₹' + subtotal + '</span></div>' +
    (packingCharge > 0 ? '<div class="pos-summary-row"><span>Packing Charges</span><span>₹' + packingCharge + '</span></div>' : '') +
    (posOrderType === 'Delivery' ? '<div class="pos-summary-row"><span>Delivery Charges</span><span>₹' + posDeliveryCharge + '</span></div>' : '') +
    (discountAmount > 0 ? '<div class="pos-summary-row"><span>Discount (' + posDiscountPercent + '%)</span><span>−₹' + discountAmount + '</span></div>' : '') +
    '<div class="pos-summary-row total"><span>Total</span><span>₹' + grandTotal + '</span></div>' +
    '<button class="pos-details-toggle" onclick="togglePosDetails()">' + (posDetailsExpanded ? '▲ Hide' : '▼') + ' Customer Details (' + detailsPreview + ')</button>' +
    (posDetailsExpanded ? '<div class="pos-details-panel">' + detailsFields + '</div>' : '') +
    '<div class="pos-action-row">' +
      '<button class="pos-btn-secondary" onclick="posUndoLastItem()">↩ Undo</button>' +
      '<button class="pos-btn-secondary" onclick="posHoldOrder()">⏸ Hold</button>' +
      '<button class="pos-btn-secondary" onclick="posClearCart()">✕ Clear</button>' +
    '</div>' +
    '<button class="pos-btn-confirm" id="posConfirmBtn" onclick="posConfirmOrder()" ' + (items.length === 0 ? 'disabled' : '') + '>✅ Confirm & Print</button>';
}

function togglePosDetails() {
  captureCartFields();
  posDetailsExpanded = !posDetailsExpanded;
  renderPosCart();
}

function posSetDeliveryCharge(amount) {
  posDeliveryCharge = amount;
  captureCartFields();
  renderPosCart();
}

function posFocusDeliveryCharge() {
  const el = document.getElementById('posDeliveryCharge');
  if (el) el.focus();
}

let posCustomerName = '', posCustomerPhone = '', posAddress = '', posNotes = '';

function setPosOrderType(t) {
  posOrderType = t;
  renderPosOrderTypeBar();
  renderPosCart();
}

// Capture field values before any re-render wipes the DOM inputs.
function captureCartFields() {
  const nameEl = document.getElementById('posName');
  const phoneEl = document.getElementById('posPhone');
  const addrEl = document.getElementById('posAddress');
  const notesEl = document.getElementById('posNotes');
  if (nameEl) posCustomerName = nameEl.value;
  if (phoneEl) posCustomerPhone = phoneEl.value;
  if (addrEl) posAddress = addrEl.value;
  if (notesEl) posNotes = notesEl.value;
}

// ── POS: hold / resume orders ───────────────────────────────────────
function posHoldOrder() {
  if (Object.keys(posCart).length === 0) { alert('Cart is empty — nothing to hold.'); return; }
  captureCartFields();
  const suggested = posCustomerName || ('Order ' + (posHeldOrders.length + 1));
  const custom = prompt('Label for this held order (e.g. a table or name):', suggested);
  if (custom === null) return; // cancelled
  const label = custom.trim() || suggested;
  posHeldOrders.push({
    label, cart: posCart, orderType: posOrderType,
    name: posCustomerName, phone: posCustomerPhone, address: posAddress, notes: posNotes,
    discountPercent: posDiscountPercent, deliveryCharge: posDeliveryCharge
  });
  posCart = {}; posCustomerName = ''; posCustomerPhone = ''; posAddress = ''; posNotes = '';
  posDiscountPercent = 0; posDeliveryCharge = 0; posOrderType = 'Dine In'; posDetailsExpanded = false;
  renderPOS();
}

function resumeHeldOrder(index) {
  if (Object.keys(posCart).length > 0) {
    if (!confirm('This will replace your current in-progress order. Continue?')) return;
  }
  const held = posHeldOrders.splice(index, 1)[0];
  posCart = held.cart;
  posOrderType = held.orderType;
  posCustomerName = held.name; posCustomerPhone = held.phone; posAddress = held.address; posNotes = held.notes;
  posDiscountPercent = held.discountPercent || 0; posDeliveryCharge = held.deliveryCharge || 0;
  posDetailsExpanded = false;
  renderPOS();
}

// ── POS: confirm & print (skips the pending queue entirely) ─────────
function posConfirmOrder() {
  captureCartFields();

  if (posOrderType === 'Delivery') {
    if (!posCustomerPhone.trim()) { alert('Phone number is required for Delivery orders.'); return; }
    if (!posAddress.trim()) { alert('Delivery address is required.'); return; }
  }
  if (Object.keys(posCart).length === 0) { alert('Cart is empty.'); return; }

  const btn = document.getElementById('posConfirmBtn');
  btn.disabled = true;
  btn.textContent = 'Sending to printer...';

  const items = Object.values(posCart).map(i => ({ id: i.id, qty: i.qty, note: i.note || '' }));

  fetch('/api/pos-order', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      order_type: posOrderType,
      customer_name: posCustomerName,
      customer_phone: posCustomerPhone,
      address: posAddress,
      notes: posNotes,
      items: items,
      discount_percent: posDiscountPercent,
      delivery_charge: posDeliveryCharge
    })
  }).then(r => r.json()).then(body => {
    if (body.status === 'ok' || body.status === 'print_failed') {
      if (body.status === 'print_failed') alert('Order #' + body.order_number + ' saved, but printing failed — check the printer.');
      posCart = {}; posCustomerName = ''; posCustomerPhone = ''; posAddress = ''; posNotes = '';
      posDiscountPercent = 0; posDeliveryCharge = 0; posOrderType = 'Dine In'; posDetailsExpanded = false;
      renderPOS();
    } else {
      alert('Could not place order: ' + (body.message || 'Unknown error'));
      btn.disabled = false;
      btn.textContent = '✅ Confirm & Print';
    }
  }).catch(e => {
    alert('Network error — could not reach the server. Try again.');
    btn.disabled = false;
    btn.textContent = '✅ Confirm & Print';
  });
}
// ─────────────────────────────────────────────────────────────────────

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
    app.run(host="0.0.0.0", port=5000, debug=False)

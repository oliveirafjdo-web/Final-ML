
import sqlite3
from pathlib import Path

DB_PATH = Path("vendas.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            sku TEXT,
            variable_cost REAL NOT NULL DEFAULT 0,
            default_price REAL NOT NULL DEFAULT 0
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            quantity REAL NOT NULL,
            unit_price REAL NOT NULL,
            marketplace_fee REAL DEFAULT 0,
            other_variable_cost REAL DEFAULT 0,
            discount REAL DEFAULT 0,
            cost_unit_at_sale REAL DEFAULT 0,
            source TEXT,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS stock_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            quantity REAL NOT NULL,
            cost_unit REAL NOT NULL,
            date TEXT NOT NULL,
            origin TEXT,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS product_inventory (
            product_id INTEGER PRIMARY KEY,
            quantity REAL NOT NULL DEFAULT 0,
            avg_cost REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)

    def ensure_setting(key, value):
        cur.execute("SELECT value FROM settings WHERE key = ?", (key,))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO settings (key, value) VALUES (?, ?)", (key, value))

    ensure_setting("imposto_pct", "0.05")
    ensure_setting("despesa_pct", "0.035")

    conn.commit()
    conn.close()

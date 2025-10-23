import os
import csv
import json
import sqlite3
from pathlib import Path
from contextlib import contextmanager
from datetime import datetime
from typing import Optional, Any, Callable, Generator, cast
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import codecs
import csv
from pathlib import Path
PANDAS_OK = True
try:
    import pandas as pd
except ImportError:
    PANDAS_OK = False
REPORTLAB_OK = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.units import mm
    from reportlab.lib.utils import simpleSplit
except Exception:
    REPORTLAB_OK = False
REPORTLAB_OK = True
PANDAS_OK = PANDAS_OK  # keep your existing flag
SMART_QUOTES_MAP = {
    "\u2018": "'", "\u2019": "'",   # left/right single
    "\u201C": '"', "\u201D": '"',   # left/right double
    "\u2013": "-",  "\u2014": "-",  # en/em dash
    "\u00A0": " ",                  # nbsp
}
DB_PATH = Path("inventory.db")
SUPPLIER_COLS_CONFIG = {
    "id": {"width": 60, "anchor": "w"},
    "name": {"width": 220, "anchor": "w"},
    "email": {"width": 240, "anchor": "w"},
    "phone": {"width": 160, "anchor": "w"},
}
SUPPLIER_FILE_COLS = ["name", "email", "phone"]
PRODUCT_COLS_CONFIG = {
    "id": {"width": 60, "anchor": "w"},
    "sku": {"width": 140, "anchor": "w"},
    "name": {"width": 260, "anchor": "w"},
    "quantity": {"width": 80, "anchor": "e"},
    "price": {"width": 90, "anchor": "e"},
    "supplier": {"width": 180, "anchor": "w"},
    "pack_size": {"width": 110, "anchor": "center"},
}
PRODUCT_FILE_COLS = ["sku", "name", "quantity", "price", "supplier", "pack_size"]
SUPPLIER_DETAIL_PRODUCT_COLS_CONFIG = {
    "id": {"width": 60, "anchor": "w"},
    "sku": {"width": 140, "anchor": "w"},
    "name": {"width": 260, "anchor": "w"},
    "pack": {"width": 100, "anchor": "center"},
    "quantity": {"width": 80, "anchor": "e"},
    "price": {"width": 90, "anchor": "e"},
}
ORDER_COLS_CONFIG = {
    "id": {"width": 60, "anchor": "w"},
    "customer": {"width": 220, "anchor": "w"},
    "date": {"width": 150, "anchor": "w"},
    "status": {"width": 110, "anchor": "w"},
    "total": {"width": 110, "anchor": "e"},
}
ORDER_STATUSES = ["Pending", "Completed", "Cancelled"]
PACK_SIZES = ["small", "medium", "large", "bulky", "extra_bulky"]
SETTINGS_DEFAULTS: dict[str, Any] = {
    "supplier_name_unique": False,  # UI-only; DB still has UNIQUE, so leave DB as-is (will show error if dup)
    "require_supplier_id_on_product": True,
    "allow_update_product_supplier": True,
    "allow_update_product_price": True,
    "allow_update_product_name": True,
    "enforce_price_gt_zero_on_import": True,
    "enforce_quantity_ge_zero_on_import": True,
    "freight_rate_small": 5.0,
    "freight_rate_medium": 10.0,
    "freight_rate_large": 20.0,
    "freight_rate_bulky": 35.0,
    "freight_rate_extra_bulky": 60.0,
    "freight_default_percent_of_price": 10.0,
}
def db_connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA foreign_keys = ON;")
    return con
def db_init() -> None:
    with db_connect() as con:
        cur = con.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                email TEXT,
                phone TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sku TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 0,
                price REAL NOT NULL DEFAULT 0,
                supplier_id INTEGER,
                pack_size TEXT,
                FOREIGN KEY (supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_name TEXT NOT NULL,
                order_date TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'Pending'
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                quantity INTEGER NOT NULL,
                price_at_order REAL NOT NULL,
                FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE,
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT
            )
        """)
        _maybe_add_column(con, "products", "pack_size", "TEXT")
        for k, v in SETTINGS_DEFAULTS.items():
            cur.execute("INSERT OR IGNORE INTO app_settings(key, value) VALUES (?,?)", (k, str(v)))
        con.commit()
def _maybe_add_column(con: sqlite3.Connection, table: str, col: str, decl: str):
    cols = [r[1] for r in con.execute(f"PRAGMA table_info({table})").fetchall()]
    if col not in cols:
        con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")
@contextmanager
def tx() -> Generator[sqlite3.Connection, None, None]:
    con = db_connect()
    try:
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()
def settings_get_raw(key: str) -> Optional[str]:
    with db_connect() as con:
        row = con.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
        return row[0] if row else None
def settings_set_raw(key: str, value: str) -> None:
    with db_connect() as con:
        con.execute("INSERT OR REPLACE INTO app_settings(key, value) VALUES(?,?)", (key, value))
def settings_get_bool(key: str) -> bool:
    v = settings_get_raw(key)
    if v is None:
        dv = SETTINGS_DEFAULTS.get(key, False)
        settings_set_raw(key, str(dv))
        return bool(dv)
    return v.lower() in ("1", "true", "yes", "on")
def settings_set_bool(key: str, b: bool) -> None:
    settings_set_raw(key, "true" if b else "false")
def settings_get_float(key: str) -> float:
    v = settings_get_raw(key)
    if v is None:
        dv = float(SETTINGS_DEFAULTS.get(key, 0.0))
        settings_set_raw(key, str(dv))
        return dv
    try:
        return float(v)
    except Exception:
        return float(SETTINGS_DEFAULTS.get(key, 0.0))
def settings_set_float(key: str, f: float) -> None:
    settings_set_raw(key, str(float(f)))
def kpis() -> tuple[int, int, int, int]:
    with db_connect() as con:
        cur = con.cursor()
        cur.execute("""
            SELECT
                (SELECT COUNT(*) FROM products),
                (SELECT COUNT(*) FROM products WHERE quantity > 0 AND quantity <= 3),
                (SELECT COUNT(*) FROM products WHERE quantity = 0),
                (SELECT COUNT(*) FROM suppliers)
        """)
        r = cur.fetchone()
        return r[0], r[1], r[2], r[3]
def suppliers_list() -> list[tuple]:
    with db_connect() as con:
        return con.execute("SELECT id, name, email, phone FROM suppliers ORDER BY name").fetchall()
def supplier_get(supplier_id: int) -> Optional[tuple]:
    with db_connect() as con:
        return con.execute("SELECT id, name, email, phone FROM suppliers WHERE id=?", (supplier_id,)).fetchone()
def get_supplier_id_by_name(con: sqlite3.Connection, name: str) -> Optional[int]:
    row = con.execute("SELECT id FROM suppliers WHERE name=?", (name,)).fetchone()
    return row[0] if row else None
def upsert_supplier(con: sqlite3.Connection, row: dict[str, Any]) -> tuple[bool, str, Optional[int]]:
    name = (row.get("name") or "").strip()
    if not name:
        return False, "Supplier 'name' is required.", None
    email = (row.get("email") or "").strip() or None
    phone = (row.get("phone") or "").strip() or None
    ex = con.execute("SELECT id FROM suppliers WHERE name=?", (name,)).fetchone()
    if ex:
        sid = ex[0]
        con.execute("UPDATE suppliers SET email=?, phone=? WHERE id=?", (email, phone, sid))
        return True, "updated", sid
    cur = con.execute("INSERT INTO suppliers(name, email, phone) VALUES(?,?,?)", (name, email, phone))
    return True, "inserted", cur.lastrowid
def supplier_insert(name: str, email: Optional[str], phone: Optional[str]) -> None:
    with db_connect() as con:
        con.execute("INSERT INTO suppliers(name, email, phone) VALUES(?,?,?)", (name, email or None, phone or None))
def supplier_update(supplier_id: int, name: str, email: Optional[str], phone: Optional[str]) -> None:
    with db_connect() as con:
        con.execute("UPDATE suppliers SET name=?, email=?, phone=? WHERE id=?", (name, email or None, phone or None, supplier_id))
def supplier_delete(supplier_id: int) -> None:
    with db_connect() as con:
        con.execute("DELETE FROM suppliers WHERE id=?", (supplier_id,))
def products_list(search_term: Optional[str] = None) -> list[tuple]:
    sql = """
        SELECT p.id, p.sku, p.name, p.quantity, p.price, COALESCE(s.name,''), p.pack_size
        FROM products p
        LEFT JOIN suppliers s ON s.id = p.supplier_id
    """
    params: list[Any] = []
    if search_term:
        sql += " WHERE p.sku LIKE ? OR p.name LIKE ? OR s.name LIKE ?"
        t = f"%{search_term}%"
        params.extend([t, t, t])
    sql += " ORDER BY p.id"
    with db_connect() as con:
        return con.execute(sql, params).fetchall()
def products_by_supplier(supplier_id: int) -> list[tuple]:
    with db_connect() as con:
        return con.execute("""
            SELECT p.id, p.sku, p.name, p.pack_size, p.quantity, p.price
            FROM products p
            WHERE p.supplier_id=?
            ORDER BY p.id
        """, (supplier_id,)).fetchall()
def product_get_details(product_id: int) -> Optional[tuple]:
    with db_connect() as con:
        return con.execute("""
            SELECT p.id, p.sku, p.name, p.quantity, p.price, p.supplier_id, COALESCE(s.name,''), p.pack_size
            FROM products p
            LEFT JOIN suppliers s ON s.id = p.supplier_id
            WHERE p.id=?
        """, (product_id,)).fetchone()
def product_get_by_sku(con: sqlite3.Connection, sku: str) -> Optional[tuple]:
    return con.execute("SELECT id FROM products WHERE sku=?", (sku,)).fetchone()
def product_insert(sku: str, name: str, quantity: int, price: float, supplier_id: Optional[int], pack_size: Optional[str]) -> None:
    with db_connect() as con:
        con.execute("""INSERT INTO products(sku,name,quantity,price,pack_size,supplier_id)
                       VALUES(?,?,?,?,?,?)""", (sku, name, quantity, price, pack_size, supplier_id))
def product_update(product_id: int, sku: str, name: str, quantity: int, price: float, supplier_id: Optional[int], pack_size: Optional[str]) -> None:
    with db_connect() as con:
        con.execute("""UPDATE products SET sku=?, name=?, quantity=?, price=?, pack_size=?, supplier_id=?
                       WHERE id=?""", (sku, name, quantity, price, pack_size, supplier_id, product_id))
def product_delete(product_id: int) -> None:
    with db_connect() as con:
        con.execute("DELETE FROM products WHERE id=?", (product_id,))
def product_upsert(con: sqlite3.Connection,
                   sku: str, name: str, quantity: int, price: float,
                   supplier_id: Optional[int],
                   update_policy: Optional[dict] = None,
                   pack_size: Optional[str] = None) -> str:
    if update_policy is None:
        update_policy = {
            "allow_update_product_supplier": settings_get_bool("allow_update_product_supplier"),
            "allow_update_product_price": settings_get_bool("allow_update_product_price"),
            "allow_update_product_name": settings_get_bool("allow_update_product_name"),
        }
    ex = con.execute("SELECT id,name,quantity,price,supplier_id,pack_size FROM products WHERE sku=?", (sku,)).fetchone()
    if ex:
        pid, old_name, _oq, old_price, old_sup, old_pack = ex
        new_name = name if update_policy.get("allow_update_product_name", True) else old_name
        new_price = price if update_policy.get("allow_update_product_price", True) else old_price
        new_sup = supplier_id if update_policy.get("allow_update_product_supplier", True) else old_sup
        new_pack = pack_size if (pack_size in PACK_SIZES or pack_size is None) else old_pack
        con.execute("""UPDATE products SET name=?, quantity=?, price=?, pack_size=?, supplier_id=? WHERE id=?""",
                    (new_name, quantity, new_price, new_pack, new_sup, pid))
        return "updated"
    else:
        norm_pack = pack_size if pack_size in PACK_SIZES else None
        con.execute("""INSERT INTO products(sku,name,quantity,price,pack_size,supplier_id)
                       VALUES(?,?,?,?,?,?)""", (sku, name, quantity, price, norm_pack, supplier_id))
        return "inserted"
def adjust_stock(con: sqlite3.Connection, product_id: int, quantity_change: int) -> None:
    con.execute("UPDATE products SET quantity = quantity + ? WHERE id = ?", (quantity_change, product_id))
def orders_list_summary() -> list[tuple]:
    with db_connect() as con:
        return con.execute("""
            SELECT o.id, o.customer_name, o.order_date, o.status,
                   COALESCE(SUM(oi.quantity * oi.price_at_order), 0) AS total
            FROM orders o
            LEFT JOIN order_items oi ON oi.order_id = o.id
            GROUP BY o.id
            ORDER BY o.order_date DESC
        """).fetchall()
def orders_list_by_supplier(supplier_id: int) -> list[tuple]:
    with db_connect() as con:
        return con.execute("""
            SELECT o.id, o.customer_name, o.order_date, o.status,
                   COALESCE(SUM(oi.quantity * oi.price_at_order), 0) AS total
            FROM orders o
            JOIN order_items oi ON oi.order_id = o.id
            JOIN products p ON p.id = oi.product_id
            WHERE p.supplier_id = ?
            GROUP BY o.id
            ORDER BY o.order_date DESC
        """, (supplier_id,)).fetchall()
def order_get_details(order_id: int) -> Optional[tuple]:
    with db_connect() as con:
        return con.execute("SELECT id, customer_name, order_date, status FROM orders WHERE id=?", (order_id,)).fetchone()
def order_get_items(order_id: int) -> list[tuple]:
    with db_connect() as con:
        return con.execute("""
            SELECT p.sku, p.name, oi.quantity, oi.price_at_order, p.id, oi.product_id, p.pack_size
            FROM order_items oi
            JOIN products p ON p.id = oi.product_id
            WHERE oi.order_id=?
        """, (order_id,)).fetchall()
def order_create(customer_name: str, items_list: list[dict]) -> int:
    with tx() as con:
        for it in items_list:
            row = con.execute("SELECT quantity FROM products WHERE id=?", (it['product_id'],)).fetchone()
            if not row or row[0] < it['quantity']:
                raise ValueError(f"Not enough stock for {it.get('sku','?')}. Available: {row[0] if row else 0}")
        date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = con.execute("INSERT INTO orders(customer_name, order_date, status) VALUES (?,?, 'Pending')",
                          (customer_name, date_str))
        order_id = cur.lastrowid
        for it in items_list:
            con.execute("""INSERT INTO order_items(order_id, product_id, quantity, price_at_order)
                           VALUES (?,?,?,?)""", (order_id, it['product_id'], it['quantity'], it['price_at_order']))
            adjust_stock(con, it['product_id'], -it['quantity'])
        return order_id
def order_delete_and_restock(order_id: int) -> None:
    with tx() as con:
        st = con.execute("SELECT status FROM orders WHERE id=?", (order_id,)).fetchone()
        if not st: return
        if st[0] != "Cancelled":
            items = con.execute("SELECT product_id, quantity FROM order_items WHERE order_id=?", (order_id,)).fetchall()
            for pid, qty in items:
                adjust_stock(con, pid, qty)
        con.execute("DELETE FROM orders WHERE id=?", (order_id,))
def order_update_status(order_id: int, new_status: str) -> None:
    if new_status not in ORDER_STATUSES:
        raise ValueError("Invalid order status.")
    with tx() as con:
        row = con.execute("SELECT status FROM orders WHERE id=?", (order_id,)).fetchone()
        if not row: raise ValueError("Order not found.")
        old = row[0]
        if old == new_status: return
        items = con.execute("SELECT product_id, quantity FROM order_items WHERE order_id=?", (order_id,)).fetchall()
        if new_status == "Cancelled" and old != "Cancelled":
            for pid, qty in items: adjust_stock(con, pid, qty)
        elif new_status != "Cancelled" and old == "Cancelled":
            for pid, qty in items:
                r = con.execute("SELECT quantity FROM products WHERE id=?", (pid,)).fetchone()
                if not r or r[0] < qty:
                    raise ValueError(f"Not enough stock to un-cancel. Product ID {pid}.")
            for pid, qty in items:
                adjust_stock(con, pid, -qty)
        con.execute("UPDATE orders SET status=? WHERE id=?", (new_status, order_id))
def _normalize_headers(headers: list[str]) -> list[str]:
    return [h.strip().lower().replace(" ", "_") for h in headers]
def _normalize_headers(headers: list[str]) -> list[str]:
    if not headers:
        return []
    fixed = []
    for h in headers:
        if h is None:
            fixed.append("")
            continue
        s = str(h).strip()
        for bad, good in SMART_QUOTES_MAP.items():
            s = s.replace(bad, good)
        s = s.replace("\ufeff", "")  # BOM
        fixed.append(s.lower().replace(" ", "_"))
    return fixed
def _clean_str_val(v: Any) -> Any:
    if not isinstance(v, str):
        return v
    s = v.replace("\ufeff", "")  # BOM
    for bad, good in SMART_QUOTES_MAP.items():
        s = s.replace(bad, good)
    return s.strip()
def _sniff_csv_dialect(sample_bytes: bytes) -> csv.Dialect:
    sample = sample_bytes.decode("latin1", errors="ignore")
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect
    except Exception:
        class _Default(csv.Dialect):
            delimiter = ","
            quotechar = '"'
            escapechar = None
            doublequote = True
            skipinitialspace = True
            lineterminator = "\n"
            quoting = csv.QUOTE_MINIMAL
        return _Default
def _try_read_csv(path: os.PathLike | str, enc: str) -> list[dict[str, Any]]:
    with open(path, "rb") as fb:
        sample = fb.read(4096)
    dialect = _sniff_csv_dialect(sample)
    rows: list[dict[str, Any]] = []
    with open(path, "r", encoding=enc, errors="replace", newline="") as f:
        reader = csv.DictReader(f, dialect=dialect)
        reader.fieldnames = _normalize_headers(reader.fieldnames or [])
        for raw in reader:
            row: dict[str, Any] = {}
            for k, v in raw.items():
                k2 = (k or "").strip().lower()
                row[k2] = _clean_str_val(v)
            rows.append(row)
    return rows
def read_table_from_file(path: os.PathLike | str) -> list[dict[str, Any]]:
    """
    Robust CSV/Excel reader.
    - CSV: tries encodings utf-8-sig, utf-8, cp1252, latin1; sniffs delimiter; normalizes headers.
    - XLSX/XLS (if pandas): uses dtype=str, normalizes headers, returns list[dict].
    Raises RuntimeError with actionable hint if all attempts fail.
    """
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        tried = []
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                rows = _try_read_csv(path, enc)
                return rows
            except Exception as e:
                tried.append(f"{enc}: {e}")
                continue
        raise RuntimeError(
            "CSV import failed. Tried encodings: utf-8-sig, utf-8, cp1252, latin1.\n"
            "Tip: re-save the file as UTF-8 (with BOM) or CSV (Comma delimited) in Excel.\n\n"
            + "\n".join(tried[:4])
        )
    if ext in (".xlsx", ".xls"):
        if not PANDAS_OK:
            raise RuntimeError(
                "Excel import requires pandas. Install with: pip install pandas openpyxl\n"
                "Or export your file as CSV."
            )
        try:
            engine = "openpyxl" if ext == ".xlsx" else None
            df = pd.read_excel(path, dtype=str, engine=engine)
        except ImportError as e:
            raise RuntimeError(
                f"Missing Excel engine: {e}\nInstall openpyxl: pip install openpyxl\n"
                "Or open file in Excel and save as CSV."
            )
        except Exception as e:
            raise RuntimeError(f"Excel read failed: {e}")
        df.columns = _normalize_headers(list(df.columns))
        df = df.fillna("")
        for col in df.columns:
            df[col] = df[col].map(_clean_str_val)
        return df.to_dict(orient="records")
    raise RuntimeError("Unsupported file type. Use .csv, or Excel (.xlsx/.xls) if pandas is available.")
def write_table_to_file(path: os.PathLike | str, rows: list[dict[str, Any]], columns: list[str]) -> None:
    """
    Robust writer:
    - CSV: write UTF-8 BOM (utf-8-sig) so Excel opens cleanly; normalize headers/values.
    - XLSX (if pandas): write via openpyxl.
    """
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            header = _normalize_headers(columns)
            w = csv.DictWriter(f, fieldnames=header, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                norm = {header[i]: _clean_str_val(r.get(columns[i], "")) for i in range(len(columns))}
                w.writerow(norm)
        return
    if ext == ".xlsx":
        if not PANDAS_OK:
            raise RuntimeError("Excel export requires pandas (and openpyxl). Install: pip install pandas openpyxl")
        try:
            header = _normalize_headers(columns)
            df = pd.DataFrame(
                [{header[i]: _clean_str_val(row.get(columns[i], "")) for i in range(len(columns))} for row in rows],
                columns=header
            )
            df.to_excel(path, index=False)
            return
        except Exception as e:
            raise RuntimeError(f"Excel write failed: {e}")
    raise RuntimeError("Unsupported export type. Use .csv, or .xlsx if pandas is installed.")
class TreeviewSorter:
    def __init__(self, tree: ttk.Treeview):
        self.tree = tree
        self.sort_state: dict[str, bool] = {}
        for col_id in tree["columns"]:
            tree.heading(col_id, command=lambda _c=col_id: self._sort_column(_c))
    def _sort_column(self, col_name: str):
        is_desc = self.sort_state.get(col_name, False)
        try:
            items = [(self.tree.set(iid, col_name), iid) for iid in self.tree.get_children("")]
        except tk.TclError:
            return
        try:
            if col_name == "total" or col_name == "price":
                items.sort(key=lambda x: float(str(x[0]).replace("$", "").replace(",", "")), reverse=is_desc)
            else:
                items.sort(key=lambda x: float(x[0]), reverse=is_desc)
        except ValueError:
            items.sort(key=lambda x: str(x[0]).lower(), reverse=is_desc)
        for i, (_v, iid) in enumerate(items):
            self.tree.move(iid, "", i)
        self.sort_state[col_name] = not is_desc
def create_treeview_with_scrollbar(parent: tk.Widget, columns_config: dict[str, dict]) -> ttk.Treeview:
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True)
    cols = list(columns_config.keys())
    tree = ttk.Treeview(frame, columns=cols, show="headings")
    for col, cfg in columns_config.items():
        tree.heading(col, text=col.title())
        tree.column(col, width=cfg.get("width", 100), anchor=cfg.get("anchor", "w"))
    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscroll=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    TreeviewSorter(tree)
    return tree
def compute_unit_shipping_cost(price: float, pack_size: Optional[str]) -> float:
    if pack_size in PACK_SIZES:
        key = f"freight_rate_{pack_size}"
        rate = settings_get_float(key)
        if rate and rate > 0:
            return float(rate)
    pct = settings_get_float("freight_default_percent_of_price")
    try:
        return max(0.0, float(price) * float(pct) / 100.0)
    except Exception:
        return 0.0
def get_order_summary_for_invoice(order_id: int)-> dict:
    hdr = order_get_details(order_id)
    if not hdr:
        raise ValueError("Order not found.")
    _id, customer_name, order_date, status = hdr
    items = order_get_items(order_id)
    lines = []
    goods_total, shipping_total = 0.0, 0.0
    for sku, name, qty, price, *_rest in items:
        pack_size = _rest[-1] if _rest else None
        line_total = qty * price
        goods_total += line_total
        unit_ship = compute_unit_shipping_cost(price, pack_size)
        ship_line = unit_ship * qty
        shipping_total += ship_line
        lines.append({
            "sku": sku,
            "name": name,
            "pack_size": pack_size or "",
            "qty": qty,
            "unit_price": float(price),
            "unit_shipping": float(unit_ship),
            "line_total": float(line_total),
            "shipping_line_total": float(ship_line),
        })
    return {
        "order_id": order_id,
        "customer_name": customer_name,
        "order_date": order_date,
        "status": status,
        "items": lines,
        "subtotal_goods": float(goods_total),
        "subtotal_shipping": float(shipping_total),
        "grand_total": float(goods_total + shipping_total),
    }
def generate_invoice_pdf(order_id: int, path: str) -> None:
    data = get_order_summary_for_invoice(order_id)
    if not REPORTLAB_OK:
        raise RuntimeError("PDF generation requires reportlab. Install with: pip install reportlab")
    c = rl_canvas.Canvas(path, pagesize=A4)
    W, H = A4
    x_margin, y_margin = 18*mm, 18*mm
    cursor_y = H - y_margin
    def draw_text(x, y, text, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        for line in simpleSplit(str(text), "Helvetica-Bold" if bold else "Helvetica", 10, W - 2*x_margin):
            c.drawString(x, y, line)
            y -= 12
        return y
    cursor_y = draw_text(x_margin, cursor_y, "INVOICE", bold=True) - 6
    cursor_y = draw_text(x_margin, cursor_y, f"Order #{data['order_id']}  |  Date: {data['order_date']}") - 6
    cursor_y = draw_text(x_margin, cursor_y, f"Bill To: {data['customer_name']}") - 10
    c.setLineWidth(0.6)
    c.line(x_margin, cursor_y, W - x_margin, cursor_y); cursor_y -= 14
    cols = ["SKU", "Name", "Pack", "Qty", "Unit", "Ship/Unit", "Line", "Ship Line"]
    col_x = [x_margin, x_margin+28*mm, x_margin+80*mm, x_margin+110*mm, x_margin+125*mm, x_margin+150*mm, x_margin+175*mm, x_margin+200*mm]
    for i, h in enumerate(cols):
        draw_text(col_x[i], cursor_y, h, bold=True)
    cursor_y -= 10
    c.line(x_margin, cursor_y, W - x_margin, cursor_y); cursor_y -= 8
    for it in data["items"]:
        if cursor_y < 40*mm:
            c.showPage()
            cursor_y = H - y_margin
        draw_text(col_x[0], cursor_y, it["sku"])
        draw_text(col_x[1], cursor_y, it["name"])
        draw_text(col_x[2], cursor_y, it["pack_size"])
        draw_text(col_x[3], cursor_y, str(it["qty"]))
        draw_text(col_x[4], cursor_y, f"${it['unit_price']:.2f}")
        draw_text(col_x[5], cursor_y, f"${it['unit_shipping']:.2f}")
        draw_text(col_x[6], cursor_y, f"${it['line_total']:.2f}")
        draw_text(col_x[7], cursor_y, f"${it['shipping_line_total']:.2f}")
        cursor_y -= 14
    cursor_y -= 6
    c.line(x_margin, cursor_y, W - x_margin, cursor_y); cursor_y -= 10
    draw_text(col_x[6], cursor_y, "SUBTOTAL:", bold=True)
    draw_text(col_x[7], cursor_y, f"${data['subtotal_goods']:.2f}", bold=True); cursor_y -= 10
    draw_text(col_x[6], cursor_y, "SHIPPING:", bold=True)
    draw_text(col_x[7], cursor_y, f"${data['subtotal_shipping']:.2f}", bold=True); cursor_y -= 10
    c.line(x_margin, cursor_y, W - x_margin, cursor_y); cursor_y -= 12
    draw_text(col_x[6], cursor_y, "TOTAL:", bold=True)
    draw_text(col_x[7], cursor_y, f"${data['grand_total']:.2f}", bold=True)
    c.showPage()
    c.save()
class SupplierEditDialog(tk.Toplevel):
    def __init__(self, parent, supplier_id: Optional[int], on_save: Callable[[], None]):
        super().__init__(parent)
        self.supplier_id = supplier_id
        self.on_save = on_save
        self.title("Edit Supplier" if supplier_id else "Add Supplier")
        self.resizable(False, False)
        self.name_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.phone_var = tk.StringVar()
        frm = ttk.Frame(self, padx=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Name:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Email:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.email_var, width=40).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Phone:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.phone_var, width=40).grid(row=2, column=1, sticky="w", padx=5)
        btn = ttk.Frame(frm); btn.grid(row=3, column=0, columnspan=2, pady=(10,0))
        ttk.Button(btn, text="Save", command=self._on_save).pack(side="left", padx=5)
        ttk.Button(btn, text="Cancel", command=self.destroy).pack(side="left", padx=5)
        if self.supplier_id: self._load_data()
        self.grab_set(); self.focus_set()
    def _load_data(self):
        d = supplier_get(cast(int, self.supplier_id))
        if d:
            _id, name, email, phone = d
            self.name_var.set(name)
            self.email_var.set(email or "")
            self.phone_var.set(phone or "")
    def _on_save(self):
        name = self.name_var.get().strip()
        email = self.email_var.get().strip()
        phone = self.phone_var.get().strip()
        if not name:
            messagebox.showerror("Validation Error", "Name is required.", parent=self); return
        try:
            if self.supplier_id:
                supplier_update(self.supplier_id, name, email, phone)
            else:
                supplier_insert(name, email, phone)
            self.on_save(); self.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Database Error", f"A supplier with the name '{name}' already exists.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}", parent=self)
class ProductEditDialog(tk.Toplevel):
    def __init__(self, parent, product_id: Optional[int], on_save: Callable[[], None]):
        super().__init__(parent)
        self.product_id = product_id
        self.on_save = on_save
        self.title("Edit Product" if product_id else "Add Product")
        self.resizable(False, False)
        self.sku_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.qty_var = tk.IntVar(value=0)
        self.price_var = tk.DoubleVar(value=0.0)
        self.supplier_var = tk.StringVar()
        self.pack_var = tk.StringVar()
        self.supplier_map: dict[str, Optional[int]] = {"": None}
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="SKU:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.sku_var, width=40).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Name:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.name_var, width=40).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Quantity:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.qty_var, width=40).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Price:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.price_var, width=40).grid(row=3, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Supplier:").grid(row=4, column=0, sticky="w", pady=5)
        self.supplier_combo = ttk.Combobox(frm, textvariable=self.supplier_var, width=38, state="readonly")
        self.supplier_combo.grid(row=4, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="Pack Size:").grid(row=5, column=0, sticky="w", pady=5)
        self.pack_combo = ttk.Combobox(frm, textvariable=self.pack_var, width=38, state="readonly",
                                       values=[""] + PACK_SIZES)
        self.pack_combo.grid(row=5, column=1, sticky="w", padx=5)
        btn = ttk.Frame(frm); btn.grid(row=6, column=0, columnspan=2, pady=(10,0))
        ttk.Button(btn, text="Save", command=self._on_save).pack(side="left", padx=5)
        ttk.Button(btn, text="Cancel", command=self.destroy).pack(side="left", padx=5)
        self._load_suppliers()
        if self.product_id: self._load_data()
        self.grab_set()
        self.focus_set()
    def _load_suppliers(self):
        self.supplier_map = {"": None}
        vals = [""]
        for sid, name, _e, _p in suppliers_list():
            self.supplier_map[name] = sid
            vals.append(name)
        self.supplier_combo["values"] = vals
    def _load_data(self):
        d = product_get_details(cast(int, self.product_id))
        if d:
            _id, sku, name, qty, price, _sid, supplier_name, pack_size = d
            self.sku_var.set(sku)
            self.name_var.set(name)
            self.qty_var.set(qty)
            self.price_var.set(price)
            self.supplier_var.set(supplier_name or "")
            self.pack_var.set(pack_size or "")
    def _on_save(self):
        sku = self.sku_var.get().strip()
        name = self.name_var.get().strip()
        try:
            qty = self.qty_var.get()
        except tk.TclError:
            messagebox.showerror("Validation Error", "Quantity must be an integer.", parent=self); return
        try:
            price = self.price_var.get()
        except tk.TclError:
            messagebox.showerror("Validation Error", "Price must be a number.", parent=self); return
        supplier_name = self.supplier_var.get()
        supplier_id = self.supplier_map.get(supplier_name)
        pack_size = self.pack_var.get().strip() or None
        if pack_size and pack_size not in PACK_SIZES:
            messagebox.showerror("Validation Error", "Invalid pack size.", parent=self); return
        if settings_get_bool("require_supplier_id_on_product") and supplier_id is None:
            if not messagebox.askyesno("Confirm", "No supplier assigned. Continue?", parent=self):
                return
        if not sku or not name:
            messagebox.showerror("Validation Error", "SKU and Name are required.", parent=self); return
        try:
            if self.product_id:
                product_update(self.product_id, sku, name, qty, price, supplier_id, pack_size)
            else:
                product_insert(sku, name, qty, price, supplier_id, pack_size)
            self.on_save(); self.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Database Error", f"A product with the SKU '{sku}' already exists.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}", parent=self)
class AddProductToOrderDialog(tk.Toplevel):
    def __init__(self, parent, on_add_item: Callable[[dict], None]):
        super().__init__(parent)
        self.on_add_item = on_add_item
        self.title("Add Product to Order")
        self.resizable(False, False)
        self.products_data = products_list()
        self.product_map = { f"{r[1]} - {r[2]}": r for r in self.products_data if r[3] > 0 }
        self.product_var = tk.StringVar()
        self.qty_var = tk.IntVar(value=1)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Product:").grid(row=0, column=0, sticky="w", pady=5)
        self.prod_combo = ttk.Combobox(frm, textvariable=self.product_var, values=list(self.product_map.keys()),
                                       state="readonly", width=40)
        self.prod_combo.grid(row=0, column=1, columnspan=2, sticky="w", padx=5)
        self.prod_combo.bind("<<ComboboxSelected>>", self._on_product_select)
        self.info_frame = ttk.Frame(frm); self.info_frame.grid(row=1, column=1, columnspan=2, sticky="w", padx=5, pady=2)
        self.lbl_available = ttk.Label(self.info_frame, text="Available: -"); self.lbl_available.pack(side="left")
        self.lbl_price = ttk.Label(self.info_frame, text="Price: -"); self.lbl_price.pack(side="left", padx=10)
        self.lbl_pack = ttk.Label(self.info_frame, text="Pack: -"); self.lbl_pack.pack(side="left", padx=10)
        ttk.Label(frm, text="Quantity:").grid(row=2, column=0, sticky="w", pady=5)
        self.spin = ttk.Spinbox(frm, textvariable=self.qty_var, from_=1, to=9999, width=10)
        self.spin.grid(row=2, column=1, sticky="w", padx=5)
        btn_frm = ttk.Frame(frm); btn_frm.grid(row=3, column=0, columnspan=3, pady=(10,0))
        ttk.Button(btn_frm, text="Add", command=self._on_add).pack(side="left", padx=5)
        ttk.Button(btn_frm, text="Cancel", command=self.destroy).pack(side="left", padx=5)
        self.grab_set(); self.focus_set()
        self.selected_product: Optional[tuple] = None
    def _on_product_select(self, _e=None):
        key = self.product_var.get()
        prod = self.product_map.get(key)
        if prod:
            self.selected_product = prod
            self.lbl_available.config(text=f"Available: {prod[3]}")  # quantity
            self.lbl_price.config(text=f"Price: ${float(prod[4]):.2f}")
            self.lbl_pack.config(text=f"Pack: {prod[6] or ''}")
            self.qty_var.set(1)
            try:
                self.spin.config(to=prod[3])
            except Exception:
                pass
    def _on_add(self):
        if not self.selected_product:
            messagebox.showerror("Error", "Please select a product.", parent=self); return
        try:
            qty = self.qty_var.get()
        except tk.TclError:
            messagebox.showerror("Error", "Quantity must be a valid number.", parent=self); return
        if qty <= 0:
            messagebox.showerror("Error", "Quantity must be greater than 0.", parent=self); return
        if qty > self.selected_product[3]:
            messagebox.showerror("Error", f"Only {self.selected_product[3]} available in stock.", parent=self); return
        item = {
            "product_id": self.selected_product[0],
            "sku": self.selected_product[1],
            "name": self.selected_product[2],
            "quantity": qty,
            "price_at_order": float(self.selected_product[4])
        }
        self.on_add_item(item); self.destroy()
class OrderCreateDialog(tk.Toplevel):
    def __init__(self, parent, on_save: Callable[[], None]):
        super().__init__(parent)
        self.on_save = on_save
        self.title("Create New Order")
        self.resizable(True, False); self.geometry("680x420")
        self.items_to_add: list[dict] = []
        self.customer_var = tk.StringVar()
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text="Customer:").pack(side="left")
        ttk.Entry(top, textvariable=self.customer_var, width=40).pack(side="left", fill="x", expand=True, padx=5)
        mid = ttk.Frame(self, padding=(10,0,10,10)); mid.pack(fill="both", expand=True)
        item_cols = {
            "sku": {"width": 120, "anchor": "w"},
            "name": {"width": 220, "anchor": "w"},
            "qty": {"width": 60, "anchor": "e"},
            "price": {"width": 80, "anchor": "e"},
            "line_total": {"width": 90, "anchor": "e"},
        }
        self.items_tree = create_treeview_with_scrollbar(mid, item_cols)
        self.lbl_total = ttk.Label(mid, text="Order Total: $0.00", font=("Segoe UI", 10, "bold"))
        self.lbl_total.pack(anchor="e", pady=5)
        btn = ttk.Frame(self); btn.pack(fill="x", padx=10)
        ttk.Button(btn, text="Add Product", command=self._add_product).pack(side="left")
        ttk.Button(btn, text="Remove Selected", command=self._remove).pack(side="left", padx=5)
        ttk.Button(btn, text="Create Order", command=self._create, style="Accent.TButton").pack(side="right")
        ttk.Button(btn, text="Cancel", command=self.destroy).pack(side="right", padx=5)
        self.grab_set(); self.focus_set()
    def _add_product(self):
        AddProductToOrderDialog(self, on_add_item=self._on_item_added)
    def _on_item_added(self, item: dict):
        for i, it in enumerate(self.items_to_add):
            if it["product_id"] == item["product_id"]:
                self.items_to_add[i]["quantity"] += item["quantity"]
                self._refresh()
                return
        self.items_to_add.append(item); self._refresh()
    def _remove(self):
        sel = self.items_tree.selection()
        if not sel: return
        sku = self.items_tree.item(sel[0], "values")[0]
        self.items_to_add = [it for it in self.items_to_add if it["sku"] != sku]
        self._refresh()
    def _refresh(self):
        self.items_tree.delete(*self.items_tree.get_children())
        tot = 0.0
        for it in self.items_to_add:
            line = it["quantity"] * it["price_at_order"]
            tot += line
            self.items_tree.insert("", "end", values=(it["sku"], it["name"], it["quantity"],
                                                      f"${it['price_at_order']:.2f}", f"${line:.2f}"))
        self.lbl_total.config(text=f"Order Total: ${tot:.2f}")
    def _create(self):
        customer = self.customer_var.get().strip()
        if not customer:
            messagebox.showerror("Error", "Customer name is required.", parent=self); return
        if not self.items_to_add:
            messagebox.showerror("Error", "Order must contain at least one product.", parent=self); return
        try:
            oid = order_create(customer, self.items_to_add)
            messagebox.showinfo("Success", f"Order #{oid} created. Stock updated.")
            self.on_save(); self.destroy()
        except ValueError as e:
            messagebox.showerror("Stock Error", str(e), parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Could not create order:\n{e}", parent=self)
class OrderViewDialog(tk.Toplevel):
    def __init__(self, parent, order_id: int, on_save: Callable[[], None]):
        super().__init__(parent)
        self.order_id = order_id
        self.on_save = on_save
        self.title(f"View/Manage Order #{order_id}")
        self.resizable(True, False); self.geometry("720x460")
        self.status_var = tk.StringVar()
        hdr = order_get_details(order_id)
        if not hdr:
            messagebox.showerror("Error", "Order not found."); self.destroy(); return
        _id, customer, date, status = hdr
        self.original_status = status
        self.status_var.set(status)
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text=f"Customer: {customer}").pack(anchor="w")
        ttk.Label(top, text=f"Date: {date}").pack(anchor="w")
        row = ttk.Frame(top); row.pack(anchor="w", pady=5)
        ttk.Label(row, text="Status:").pack(side="left")
        self.status_combo = ttk.Combobox(row, textvariable=self.status_var, values=ORDER_STATUSES, state="readonly")
        self.status_combo.pack(side="left", padx=5)
        mid = ttk.Frame(self, padding=(10,0,10,10)); mid.pack(fill="both", expand=True)
        cols = {
            "sku": {"width": 120, "anchor": "w"},
            "name": {"width": 220, "anchor": "w"},
            "pack": {"width": 90, "anchor": "center"},
            "qty": {"width": 60, "anchor": "e"},
            "price": {"width": 80, "anchor": "e"},
            "line_total": {"width": 100, "anchor": "e"},
        }
        self.items_tree = create_treeview_with_scrollbar(mid, cols)
        self.lbl_goods = ttk.Label(mid, text="Goods: $0.00", font=("Segoe UI", 10, "bold"))
        self.lbl_ship = ttk.Label(mid, text="Shipping: $0.00", font=("Segoe UI", 10, "bold"))
        self.lbl_total = ttk.Label(mid, text="Total: $0.00", font=("Segoe UI", 10, "bold"))
        for w in (self.lbl_goods, self.lbl_ship, self.lbl_total):
            w.pack(anchor="e", pady=2)
        self._load_items()
        btn = ttk.Frame(self); btn.pack(padx=10, fill="x")
        ttk.Button(btn, text="Save Status", command=self._save_status).pack(side="right")
        ttk.Button(btn, text="Invoice PDF", command=self._export_pdf).pack(side="right", padx=6)
        ttk.Button(btn, text="Close", command=self.destroy).pack(side="right", padx=6)
        self.grab_set(); self.focus_set()
    def _load_items(self):
        self.items_tree.delete(*self.items_tree.get_children())
        tot, ship = 0.0, 0.0
        for it in order_get_items(self.order_id):
            sku, name, qty, price, _pid, _opid, pack = it
            line = qty * price
            tot += line
            unit_ship = compute_unit_shipping_cost(price, pack)
            ship += unit_ship * qty
            self.items_tree.insert("", "end", values=(sku, name, pack or "", qty, f"${price:.2f}", f"${line:.2f}"))
        self.lbl_goods.config(text=f"Goods: ${tot:.2f}")
        self.lbl_ship.config(text=f"Shipping: ${ship:.2f}")
        self.lbl_total.config(text=f"Total: ${tot+ship:.2f}")
    def _save_status(self):
        new_status = self.status_var.get()
        if new_status == self.original_status:
            self.destroy(); return
        msg = f"Change order status from '{self.original_status}' to '{new_status}'?"
        if new_status == "Cancelled" and self.original_status != "Cancelled":
            msg += "\n\nThis will restock all items."
        elif new_status != "Cancelled" and self.original_status == "Cancelled":
            msg += "\n\nThis will deduct stock for all items."
        if not messagebox.askyesno("Confirm", msg, parent=self): return
        try:
            order_update_status(self.order_id, new_status)
            self.on_save(); self.destroy()
        except ValueError as e:
            messagebox.showerror("Stock Error", str(e), parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Could not update status:\n{e}", parent=self)
    def _export_pdf(self):
        path = filedialog.asksaveasfilename(title="Save Invoice PDF", defaultextension=".pdf",
                                            filetypes=[("PDF", "*.pdf")], parent=self)
        if not path: return
        try:
            generate_invoice_pdf(self.order_id, path)
            messagebox.showinfo("Invoice", f"Saved:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("Invoice", str(e), parent=self)
class SupplierDetailDialog(tk.Toplevel):
    """
    Notebook: Details / Products / Orders / Freight (global freight config editor)
    """
    def __init__(self, parent, supplier_id: int, on_any_change: Callable[[], None]):
        super().__init__(parent)
        self.title("Supplier Detail")
        self.resizable(True, True)
        self.supplier_id = supplier_id
        self.on_any_change = on_any_change
        info = supplier_get(supplier_id)
        if not info:
            messagebox.showerror("Error", "Supplier not found."); self.destroy(); return
        self.sid, self.supplier_name, self.supplier_email, self.supplier_phone = info
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_details = ttk.Frame(nb, padding=8); nb.add(self.tab_details, text="Details")
        self.tab_products = ttk.Frame(nb, padding=8); nb.add(self.tab_products, text="Products")
        self.tab_orders = ttk.Frame(nb, padding=8); nb.add(self.tab_orders, text="Orders")
        self.tab_freight = ttk.Frame(nb, padding=8); nb.add(self.tab_freight, text="Freight")  # global freight config
        self._build_tab_details()
        self._build_tab_products()
        self._build_tab_orders()
        self._build_tab_freight()
        self.geometry("1000x620")
        self.grab_set(); self.focus_set()
    def _build_tab_details(self):
        f = self.tab_details
        hdr = ttk.Frame(f); hdr.pack(fill="x", pady=(0,8))
        ttk.Label(hdr, text=f"Supplier: {self.supplier_name}", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(hdr, text=f"Email: {self.supplier_email or ''}").pack(anchor="w")
        ttk.Label(hdr, text=f"Phone: {self.supplier_phone or ''}").pack(anchor="w")
        btns = ttk.Frame(f); btns.pack(fill="x", pady=(6,0))
        ttk.Button(btns, text="Edit", command=self._edit_supplier).pack(side="left", padx=4)
        ttk.Button(btns, text="Delete", command=self._delete_supplier).pack(side="left", padx=4)
        ttk.Button(btns, text="Export Products", command=self._export_supplier_products).pack(side="left", padx=4)
        ttk.Button(btns, text="Close", command=self.destroy).pack(side="right", padx=4)
    def _repaint_details(self):
        for w in self.tab_details.winfo_children(): w.destroy()
        self._build_tab_details()
    def _edit_supplier(self):
        SupplierEditDialog(self, supplier_id=self.sid, on_save=self._after_supplier_change)
    def _delete_supplier(self):
        if not messagebox.askyesno("Confirm",
                                   f"Delete supplier '{self.supplier_name}'?\n\nProducts will be unassigned.",
                                   parent=self):
            return
        try:
            supplier_delete(self.sid)
            self.on_any_change(); self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete supplier:\n{e}", parent=self)
    def _after_supplier_change(self):
        info = supplier_get(self.supplier_id)
        if info:
            self.sid, self.supplier_name, self.supplier_email, self.supplier_phone = info
        self._repaint_details(); self._refresh_products(); self._refresh_orders()
        self.on_any_change()
    def _export_supplier_products(self):
        path = filedialog.asksaveasfilename(title="Export Supplier Products",
                                            defaultextension=".csv",
                                            filetypes=[("CSV","*.csv")] + ([("Excel","*.xlsx")] if PANDAS_OK else []),
                                            parent=self)
        if not path: return
        try:
            rows = []
            for r in products_by_supplier(self.supplier_id):
                rows.append({
                    "sku": r[1], "name": r[2], "pack_size": r[3] or "", "quantity": r[4], "price": r[5],
                    "supplier": self.supplier_name
                })
            write_table_to_file(path, rows, PRODUCT_FILE_COLS)
            messagebox.showinfo("Export", f"Exported:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("Export failed", str(e), parent=self)
    def _build_tab_products(self):
        f = self.tab_products
        f.rowconfigure(1, weight=1); f.columnconfigure(0, weight=1)
        hdr = ttk.Frame(f); hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(hdr, text=f"Products for {self.supplier_name}", font=("Segoe UI", 11, "bold")).pack(side="left")
        cont = ttk.Frame(f); cont.grid(row=1, column=0, sticky="nsew", pady=(6,6))
        self.tree_products = create_treeview_with_scrollbar(cont, SUPPLIER_DETAIL_PRODUCT_COLS_CONFIG)
        self.tree_products.bind("<Double-1>", self._open_product_edit)
        btns = ttk.Frame(f); btns.grid(row=2, column=0, sticky="w")
        ttk.Button(btns, text="Refresh", command=self._refresh_products).pack(side="left", padx=3)
        ttk.Button(btns, text="Add New", command=self._add_product_for_supplier).pack(side="left", padx=3)
        ttk.Button(btns, text="Edit Selected", command=self._edit_selected_product).pack(side="left", padx=3)
        ttk.Button(btns, text="Delete Selected", command=self._delete_selected_product).pack(side="left", padx=3)
        ttk.Button(btns, text="Import (Assign to this supplier)", command=self._import_products_for_supplier).pack(side="left", padx=10)
        ttk.Button(btns, text="Export", command=self._export_supplier_products).pack(side="left", padx=3)
        self._refresh_products()
    def _refresh_products(self):
        self.tree_products.delete(*self.tree_products.get_children())
        for r in products_by_supplier(self.supplier_id):
            self.tree_products.insert("", "end", values=(r[0], r[1], r[2], r[3] or "", r[4], f"{float(r[5]):.2f}"))
    def _selected_product_id(self) -> Optional[int]:
        sel = self.tree_products.selection()
        return int(self.tree_products.item(sel[0], "values")[0]) if sel else None
    def _add_product_for_supplier(self):
        dlg = ProductEditDialog(self, product_id=None, on_save=self._after_products_change)
        try: dlg.supplier_var.set(self.supplier_name)
        except Exception: pass
    def _edit_selected_product(self):
        pid = self._selected_product_id()
        if not pid:
            messagebox.showinfo("No Selection", "Select a product to edit.", parent=self); return
        ProductEditDialog(self, product_id=pid, on_save=self._after_products_change)
    def _delete_selected_product(self):
        pid = self._selected_product_id()
        if not pid:
            messagebox.showinfo("No Selection", "Select a product to delete.", parent=self); return
        name = self.tree_products.item(self.tree_products.selection()[0], "values")[2]
        if not messagebox.askyesno("Confirm Delete",
                                   f"Delete product '{name}'?\n\nThis fails if referenced by orders.",
                                   parent=self): return
        try:
            product_delete(pid)
            self._after_products_change()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Cannot delete product referenced by orders.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete product:\n{e}", parent=self)
    def _import_products_for_supplier(self):
        path = filedialog.askopenfilename(title=f"Import/Update Products for {self.supplier_name} (CSV/XLSX)",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx *.xls"), ("All","*.*")],
                                          parent=self)
        if not path: return
        try:
            rows = read_table_from_file(path)
            cleaned = []
            for r in rows:
                cleaned.append({
                    "sku": r.get("sku",""),
                    "name": r.get("name",""),
                    "quantity": r.get("quantity", r.get("qty", 0)),
                    "price": r.get("price", 0),
                    "pack_size": (r.get("pack_size","") or "").strip().lower(),
                })
            inserted, updated, errors = 0, 0, []
            enforce_price_pos = settings_get_bool("enforce_price_gt_zero_on_import")
            enforce_qty_nonneg = settings_get_bool("enforce_quantity_ge_zero_on_import")
            update_policy = {
                "allow_update_product_supplier": settings_get_bool("allow_update_product_supplier"),
                "allow_update_product_price": settings_get_bool("allow_update_product_price"),
                "allow_update_product_name": settings_get_bool("allow_update_product_name"),
            }
            with tx() as con:
                forced_supplier_id = get_supplier_id_by_name(con, self.supplier_name)
                if not forced_supplier_id:
                    ok, _st, new_sid = upsert_supplier(con, {"name": self.supplier_name})
                    forced_supplier_id = new_sid if ok else None
                if not forced_supplier_id:
                    raise RuntimeError("Could not resolve this supplier ID.")
                for i, row in enumerate(cleaned, 1):
                    sku = (row.get("sku") or "").strip()
                    name = (row.get("name") or "").strip()
                    if not sku or not name:
                        errors.append(f"Row {i}: 'sku' and 'name' required."); continue
                    try:
                        quantity = int(row.get("quantity", 0) or 0)
                        if enforce_qty_nonneg and quantity < 0: raise ValueError
                    except Exception:
                        errors.append(f"Row {i} (SKU {sku}): invalid 'quantity' (>=0)."); continue
                    try:
                        price = float(row.get("price", 0) or 0)
                        if enforce_price_pos and not (price > 0): raise ValueError
                        if not enforce_price_pos and price <= 0: price = 0.01
                        if not enforce_qty_nonneg and quantity < 0: quantity = 0
                    except Exception:
                        errors.append(f"Row {i} (SKU {sku}): invalid 'price' (>0)."); continue
                    ps = row.get("pack_size") or None
                    if ps and ps not in PACK_SIZES: ps = None
                    status = product_upsert(con, sku, name, quantity, price, forced_supplier_id, update_policy, pack_size=ps)
                    if status == "inserted": inserted += 1
                    else: updated += 1
            msg = f"{self.supplier_name}: Import completed.\nInserted: {inserted}\nUpdated: {updated}"
            if errors:
                msg += f"\nErrors: {len(errors)} (showing up to 20)\n\n" + "\n".join(errors[:20])
                messagebox.showwarning("Import issues", msg, parent=self)
            else:
                messagebox.showinfo("Import", msg, parent=self)
            self._after_products_change()
        except Exception as e:
            messagebox.showerror("Import failed", str(e), parent=self)
    def _open_product_edit(self, _e=None):
        sel = self.tree_products.selection()
        if not sel: return
        try:
            pid = int(self.tree_products.item(sel[0], "values")[0])
            ProductEditDialog(self, product_id=pid, on_save=self._after_products_change)
        except Exception:
            pass
    def _after_products_change(self):
        self._refresh_products(); self.on_any_change()
    def _build_tab_orders(self):
        f = self.tab_orders
        f.rowconfigure(1, weight=1); f.columnconfigure(0, weight=1)
        hdr = ttk.Frame(f); hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(hdr, text=f"Orders containing {self.supplier_name} products", font=("Segoe UI", 11, "bold")).pack(side="left")
        cont = ttk.Frame(f); cont.grid(row=1, column=0, sticky="nsew", pady=(6,6))
        self.tree_orders = create_treeview_with_scrollbar(cont, ORDER_COLS_CONFIG)
        self.tree_orders.bind("<Double-1>", self._open_order_view)
        btns = ttk.Frame(f); btns.grid(row=2, column=0, sticky="w")
        ttk.Button(btns, text="Refresh", command=self._refresh_orders).pack(side="left", padx=3)
        ttk.Button(btns, text="View/Manage", command=self._open_order_view).pack(side="left", padx=3)
        self._refresh_orders()
    def _refresh_orders(self):
        self.tree_orders.delete(*self.tree_orders.get_children())
        for r in orders_list_by_supplier(self.supplier_id):
            self.tree_orders.insert("", "end", values=(r[0], r[1], r[2], r[3], f"${float(r[4]):.2f}"))
    def _selected_order_id(self) -> Optional[int]:
        sel = self.tree_orders.selection()
        return int(self.tree_orders.item(sel[0], "values")[0]) if sel else None
    def _open_order_view(self, _e=None):
        oid = self._selected_order_id()
        if not oid: return
        OrderViewDialog(self, order_id=oid, on_save=self._after_orders_change)
    def _after_orders_change(self):
        self._refresh_orders(); self.on_any_change()
    def _build_tab_freight(self):
        f = self.tab_freight
        grid = ttk.LabelFrame(f, text="Per-Pack-Size Shipping Rates (per unit)", padding=10)
        grid.pack(fill="x", expand=False, pady=(0,8))
        self.var_rate_small = tk.DoubleVar(value=settings_get_float("freight_rate_small"))
        self.var_rate_medium = tk.DoubleVar(value=settings_get_float("freight_rate_medium"))
        self.var_rate_large = tk.DoubleVar(value=settings_get_float("freight_rate_large"))
        self.var_rate_bulky = tk.DoubleVar(value=settings_get_float("freight_rate_bulky"))
        self.var_rate_extra = tk.DoubleVar(value=settings_get_float("freight_rate_extra_bulky"))
        self.var_default_pct = tk.DoubleVar(value=settings_get_float("freight_default_percent_of_price"))
        rows = [("Small", self.var_rate_small),
                ("Medium", self.var_rate_medium),
                ("Large", self.var_rate_large),
                ("Bulky", self.var_rate_bulky),
                ("Extra Bulky", self.var_rate_extra)]
        for i, (label, var) in enumerate(rows):
            ttk.Label(grid, text=f"{label}:").grid(row=i, column=0, sticky="w", pady=3, padx=(4,6))
            ttk.Entry(grid, textvariable=var, width=12).grid(row=i, column=1, sticky="w", pady=3)
        df = ttk.LabelFrame(f, text="Fallback (if no pack size or missing rate)", padding=10)
        df.pack(fill="x", expand=False)
        ttk.Label(df, text="Default shipping as % of product price (per unit):").grid(row=0, column=0, sticky="w", padx=(4,6))
        ttk.Entry(df, textvariable=self.var_default_pct, width=8).grid(row=0, column=1, sticky="w")
        ttk.Label(df, text="%").grid(row=0, column=2, sticky="w")
        btns = ttk.Frame(f); btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Save Freight Settings", command=self._save_freight).pack(side="right", padx=6)
    def _save_freight(self):
        try:
            settings_set_float("freight_rate_small", float(self.var_rate_small.get()))
            settings_set_float("freight_rate_medium", float(self.var_rate_medium.get()))
            settings_set_float("freight_rate_large", float(self.var_rate_large.get()))
            settings_set_float("freight_rate_bulky", float(self.var_rate_bulky.get()))
            settings_set_float("freight_rate_extra_bulky", float(self.var_rate_extra.get()))
            settings_set_float("freight_default_percent_of_price", float(self.var_default_pct.get()))
            messagebox.showinfo("Freight", "Freight settings saved.", parent=self)
        except Exception as e:
            messagebox.showerror("Freight", f"Failed to save freight settings:\n{e}", parent=self)
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Inventory / Orders / Freight")
        self.geometry("1080x680")
        self.minsize(900, 560)
        s = ttk.Style(); s.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True)
        self.tab_dashboard = ttk.Frame(nb, padding=6)
        self.tab_orders = ttk.Frame(nb, padding=6)
        self.tab_products = ttk.Frame(nb, padding=6)
        self.tab_suppliers = ttk.Frame(nb, padding=6)
        self.tab_settings = ttk.Frame(nb, padding=6)
        nb.add(self.tab_dashboard, text="Dashboard")
        nb.add(self.tab_orders, text="Orders")
        nb.add(self.tab_products, text="Products")
        nb.add(self.tab_suppliers, text="Suppliers")
        nb.add(self.tab_settings, text="Settings")
        self._build_dashboard()
        self._build_orders()
        self._build_products()
        self._build_suppliers()
        self._build_settings()
        self._refresh_all()
    def _build_dashboard(self):
        f = self.tab_dashboard
        ttk.Label(f, text="Dashboard", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(10,6))
        self.lbl_total = ttk.Label(f, text="Total Products: 0"); self.lbl_total.pack(anchor="w", padx=10, pady=2)
        self.lbl_low = ttk.Label(f, text="Low Stock (<=3): 0"); self.lbl_low.pack(anchor="w", padx=10, pady=2)
        self.lbl_oos = ttk.Label(f, text="Out of Stock: 0"); self.lbl_oos.pack(anchor="w", padx=10, pady=2)
        self.lbl_sup = ttk.Label(f, text="Suppliers: 0"); self.lbl_sup.pack(anchor="w", padx=10, pady=2)
        ttk.Button(f, text="Refresh", command=self._refresh_dashboard).pack(anchor="w", padx=10, pady=8)
    def _refresh_dashboard(self):
        total, low, oos, sups = kpis()
        self.lbl_total.config(text=f"Total Products: {total}")
        self.lbl_low.config(text=f"Low Stock (<=3): {low}")
        self.lbl_oos.config(text=f"Out of Stock: {oos}")
        self.lbl_sup.config(text=f"Suppliers: {sups}")
    def _build_suppliers(self):
        f = self.tab_suppliers
        f.rowconfigure(1, weight=1); f.columnconfigure(0, weight=1)
        hdr = ttk.Frame(f); hdr.grid(row=0, column=0, sticky="ew", padx=8, pady=6)
        ttk.Label(hdr, text="Suppliers", font=("Segoe UI", 12, "bold")).pack(side="left")
        ttk.Button(hdr, text="Import", command=self._import_suppliers).pack(side="right", padx=4)
        ttk.Button(hdr, text="Export", command=self._export_suppliers).pack(side="right", padx=4)
        ttk.Button(hdr, text="Template", command=self._template_suppliers).pack(side="right", padx=4)
        cont = ttk.Frame(f); cont.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0,8))
        self.sup_tree = create_treeview_with_scrollbar(cont, SUPPLIER_COLS_CONFIG)
        self.sup_tree.bind("<Double-1>", self._open_supplier_detail)
        btm = ttk.Frame(f); btm.grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ttk.Button(btm, text="Refresh", command=self._refresh_suppliers).pack(side="left")
        ttk.Button(btm, text="Add New", command=self._add_supplier).pack(side="left", padx=5)
        ttk.Button(btm, text="Edit Selected", command=self._edit_supplier).pack(side="left", padx=5)
        ttk.Button(btm, text="Delete Selected", command=self._delete_supplier).pack(side="left", padx=5)
    def _open_supplier_detail(self, _e=None):
        sid = self._selected_supplier_id()
        if not sid: return
        SupplierDetailDialog(self, sid, on_any_change=self._refresh_all)
    def _selected_supplier_id(self) -> Optional[int]:
        sel = self.sup_tree.selection()
        return int(self.sup_tree.item(sel[0], "values")[0]) if sel else None
    def _refresh_suppliers(self):
        self.sup_tree.delete(*self.sup_tree.get_children())
        for row in suppliers_list():
            self.sup_tree.insert("", "end", values=row)
    def _add_supplier(self):
        SupplierEditDialog(self, supplier_id=None, on_save=self._refresh_all)
    def _edit_supplier(self):
        sid = self._selected_supplier_id()
        if not sid:
            messagebox.showinfo("No Selection", "Select a supplier to edit."); return
        SupplierEditDialog(self, supplier_id=sid, on_save=self._refresh_all)
    def _delete_supplier(self):
        sid = self._selected_supplier_id()
        if not sid:
            messagebox.showinfo("No Selection", "Select a supplier to delete."); return
        name = self.sup_tree.item(self.sup_tree.selection()[0], "values")[1]
        if messagebox.askyesno("Confirm Delete",
                               f"Delete '{name}'?\n\nAll products from this supplier will be unassigned."):
            try:
                supplier_delete(sid); self._refresh_all()
            except Exception as e:
                messagebox.showerror("Error", f"Could not delete supplier:\n{e}")
    def _import_suppliers(self):
        path = filedialog.askopenfilename(title="Import Suppliers (CSV/XLSX)",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx *.xls"), ("All","*.*")])
        if not path: return
        try:
            rows = read_table_from_file(path)
            cleaned = [{k: r.get(k,"") for k in SUPPLIER_FILE_COLS} for r in rows]
            inserted, updated, errors = 0, 0, []
            with tx() as con:
                for i, row in enumerate(cleaned, 1):
                    ok, status, _sid = upsert_supplier(con, row)
                    if not ok:
                        errors.append(f"Row {i}: {status}")
                    else:
                        if status == "inserted": inserted += 1
                        else: updated += 1
            msg = f"Suppliers import complete.\nInserted: {inserted}\nUpdated: {updated}"
            if errors:
                msg += f"\nErrors: {len(errors)} (up to 20):\n" + "\n".join(errors[:20])
                messagebox.showwarning("Import finished with issues", msg)
            else:
                messagebox.showinfo("Import", msg)
            self._refresh_all()
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
    def _export_suppliers(self):
        rows = [{"name": r[1], "email": r[2], "phone": r[3]} for r in suppliers_list()]
        self._export_data_file("Export Suppliers", SUPPLIER_FILE_COLS, rows)
    def _template_suppliers(self):
        self._save_template_file("Save Suppliers Template", SUPPLIER_FILE_COLS)
    def _build_products(self):
        f = self.tab_products
        f.rowconfigure(2, weight=1); f.columnconfigure(0, weight=1)
        hdr = ttk.Frame(f); hdr.grid(row=0, column=0, sticky="ew", padx=8, pady=6)
        ttk.Label(hdr, text="Products", font=("Segoe UI", 12, "bold")).pack(side="left")
        ttk.Button(hdr, text="Import", command=self._import_products).pack(side="right", padx=4)
        ttk.Button(hdr, text="Export", command=self._export_products).pack(side="right", padx=4)
        ttk.Button(hdr, text="Template", command=self._template_products).pack(side="right", padx=4)
        search = ttk.Frame(f); search.grid(row=1, column=0, sticky="ew", padx=8, pady=(0,6))
        ttk.Label(search, text="Search:").pack(side="left", padx=(0,5))
        self.prod_search_var = tk.StringVar()
        self.prod_search_entry = ttk.Entry(search, textvariable=self.prod_search_var, width=40)
        self.prod_search_entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(search, text="Find", command=self._search_products).pack(side="left", padx=5)
        ttk.Button(search, text="Clear", command=self._clear_search_products).pack(side="left", padx=5)
        self.prod_search_entry.bind("<Return>", self._search_products)
        cont = ttk.Frame(f); cont.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0,8))
        self.prod_tree = create_treeview_with_scrollbar(cont, PRODUCT_COLS_CONFIG)
        self.prod_tree.bind("<Double-1>", self._edit_product_event)
        btm = ttk.Frame(f); btm.grid(row=3, column=0, sticky="w", padx=8, pady=6)
        ttk.Button(btm, text="Refresh", command=self._refresh_products).pack(side="left")
        ttk.Button(btm, text="Add New", command=self._add_product).pack(side="left", padx=5)
        ttk.Button(btm, text="Edit Selected", command=self._edit_product).pack(side="left", padx=5)
        ttk.Button(btm, text="Delete Selected", command=self._delete_product).pack(side="left", padx=5)
    def _edit_product_event(self, _e=None): self._edit_product()
    def _selected_product_id(self) -> Optional[int]:
        sel = self.prod_tree.selection()
        return int(self.prod_tree.item(sel[0], "values")[0]) if sel else None
    def _search_products(self, _e=None):
        term = self.prod_search_var.get().strip()
        self._refresh_products(term)
    def _clear_search_products(self):
        self.prod_search_var.set(""); self._refresh_products()
    def _refresh_products(self, search_term: Optional[str] = None):
        if search_term is None: search_term = self.prod_search_var.get().strip() or None
        self.prod_tree.delete(*self.prod_tree.get_children())
        for r in products_list(search_term):
            self.prod_tree.insert("", "end", values=(
                r[0], r[1], r[2], r[3], f"{float(r[4]):.2f}", r[5], r[6] or ""
            ))
    def _add_product(self):
        ProductEditDialog(self, product_id=None, on_save=self._refresh_all)
    def _edit_product(self):
        pid = self._selected_product_id()
        if not pid:
            messagebox.showinfo("No Selection", "Select a product to edit."); return
        ProductEditDialog(self, product_id=pid, on_save=self._refresh_all)
    def _delete_product(self):
        pid = self._selected_product_id()
        if not pid:
            messagebox.showinfo("No Selection", "Select a product to delete."); return
        name = self.prod_tree.item(self.prod_tree.selection()[0], "values")[2]
        if messagebox.askyesno("Confirm Delete",
                               f"Delete product:\n\n'{name}'?\n\nThis will fail if product is referenced by orders."):
            try:
                product_delete(pid); self._refresh_all()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Cannot delete product referenced by orders.")
            except Exception as e:
                messagebox.showerror("Error", f"Could not delete product:\n{e}")
    def _import_products(self):
        path = filedialog.askopenfilename(title="Import/Update Products (CSV/XLSX)",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx *.xls"), ("All","*.*")])
        if not path: return
        try:
            rows = read_table_from_file(path)
            cleaned = [{k: r.get(k,"") for k in PRODUCT_FILE_COLS} for r in rows]
            inserted, updated, errors = 0, 0, []
            enforce_price_pos = settings_get_bool("enforce_price_gt_zero_on_import")
            enforce_qty_nonneg = settings_get_bool("enforce_quantity_ge_zero_on_import")
            update_policy = {
                "allow_update_product_supplier": settings_get_bool("allow_update_product_supplier"),
                "allow_update_product_price": settings_get_bool("allow_update_product_price"),
                "allow_update_product_name": settings_get_bool("allow_update_product_name"),
            }
            with tx() as con:
                for i, row in enumerate(cleaned, 1):
                    sku = (row.get("sku") or "").strip()
                    name = (row.get("name") or "").strip()
                    if not sku or not name:
                        errors.append(f"Row {i}: 'sku' and 'name' required."); continue
                    try:
                        quantity = int(row.get("quantity", 0) or 0)
                        if enforce_qty_nonneg and quantity < 0: raise ValueError
                    except Exception:
                        errors.append(f"Row {i} (SKU {sku}): invalid 'quantity' (>=0)."); continue
                    try:
                        price = float(row.get("price", 0) or 0)
                        if enforce_price_pos and not (price > 0): raise ValueError
                        if not enforce_price_pos and price <= 0: price = 0.01
                        if not enforce_qty_nonneg and quantity < 0: quantity = 0
                    except Exception:
                        errors.append(f"Row {i} (SKU {sku}): invalid 'price' (>0)."); continue
                    supplier_name = (row.get("supplier") or "").strip()
                    supplier_id: Optional[int] = None
                    if supplier_name:
                        supplier_id = get_supplier_id_by_name(con, supplier_name)
                        if not supplier_id:
                            ok, _st, new_sid = upsert_supplier(con, {"name": supplier_name})
                            if ok: supplier_id = new_sid
                            else:
                                errors.append(f"Row {i} (SKU {sku}): cannot create supplier '{supplier_name}'."); continue
                    elif settings_get_bool("require_supplier_id_on_product"):
                        pass
                    ps = (row.get("pack_size") or "").strip().lower()
                    if ps and ps not in PACK_SIZES: ps = None
                    status = product_upsert(con, sku, name, quantity, price, supplier_id, update_policy, pack_size=ps)
                    if status == "inserted": inserted += 1
                    else: updated += 1
            msg = f"Product import complete.\nInserted new: {inserted}\nUpdated existing: {updated}"
            if errors:
                msg += f"\nErrors: {len(errors)} (up to 20):\n" + "\n".join(errors[:20])
                messagebox.showwarning("Import completed with issues", msg)
            else:
                messagebox.showinfo("Import", msg)
            self._refresh_all()
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
    def _export_products(self):
        rows = []
        for r in products_list():
            rows.append({
                "sku": r[1], "name": r[2], "quantity": r[3], "price": r[4],
                "supplier": r[5], "pack_size": r[6] or ""
            })
        self._export_data_file("Export Products", PRODUCT_FILE_COLS, rows)
    def _template_products(self):
        self._save_template_file("Save Products Template", PRODUCT_FILE_COLS)
    def _build_orders(self):
        f = self.tab_orders
        f.rowconfigure(1, weight=1); f.columnconfigure(0, weight=1)
        hdr = ttk.Frame(f); hdr.grid(row=0, column=0, sticky="ew", padx=8, pady=6)
        ttk.Label(hdr, text="Orders", font=("Segoe UI", 12, "bold")).pack(side="left")
        ttk.Button(hdr, text="Import", command=self._import_orders).pack(side="right", padx=4)  # import with items_json
        cont = ttk.Frame(f); cont.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0,8))
        self.orders_tree = create_treeview_with_scrollbar(cont, ORDER_COLS_CONFIG)
        self.orders_tree.bind("<Double-1>", self._open_order_view)
        btm = ttk.Frame(f); btm.grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ttk.Button(btm, text="Refresh", command=self._refresh_orders).pack(side="left")
        ttk.Button(btm, text="Add New Order", command=self._add_order, style="Accent.TButton").pack(side="left", padx=5)
        ttk.Button(btm, text="View/Manage Order", command=self._open_order_view).pack(side="left", padx=5)
        ttk.Button(btm, text="Delete Order", command=self._delete_order).pack(side="left", padx=5)
    def _refresh_orders(self):
        self.orders_tree.delete(*self.orders_tree.get_children())
        for r in orders_list_summary():
            self.orders_tree.insert("", "end", values=(r[0], r[1], r[2], r[3], f"${float(r[4]):.2f}"))
    def _selected_order_id(self) -> Optional[int]:
        sel = self.orders_tree.selection()
        return int(self.orders_tree.item(sel[0], "values")[0]) if sel else None
    def _add_order(self):
        OrderCreateDialog(self, on_save=self._refresh_all)
    def _open_order_view(self, _e=None):
        oid = self._selected_order_id()
        if not oid:
            messagebox.showinfo("No Selection", "Select an order to view/manage."); return
        OrderViewDialog(self, order_id=oid, on_save=self._refresh_all)
    def _delete_order(self):
        oid = self._selected_order_id()
        if not oid:
            messagebox.showinfo("No Selection", "Select an order to delete."); return
        status = self.orders_tree.item(self.orders_tree.selection()[0], "values")[3]
        msg = f"Delete Order #{oid}?"
        if status != "Cancelled": msg += "\n\nThis will restock the items."
        if messagebox.askyesno("Confirm Delete", msg):
            try:
                order_delete_and_restock(oid); self._refresh_all()
            except Exception as e:
                messagebox.showerror("Error", f"Could not delete order:\n{e}")
    def _import_orders(self):
        """
        Import with headers: customer_name, order_date(optional), status(optional), items_json (required)
        items_json is JSON array: [{product_sku, quantity, price_at_order}]
        """
        path = filedialog.askopenfilename(title="Import Orders (CSV/XLSX)",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx *.xls"), ("All","*.*")])
        if not path: return
        try:
            rows = read_table_from_file(path)
            cleaned = []
            for r in rows:
                cleaned.append({
                    "customer_name": r.get("customer_name","").strip(),
                    "order_date": r.get("order_date","").strip(),
                    "status": r.get("status","").strip(),
                    "items_json": r.get("items_json","").strip(),
                })
            created, errors = 0, []
            with tx() as con:
                for i, row in enumerate(cleaned, 1):
                    cust = row["customer_name"]
                    if not cust:
                        errors.append(f"Row {i}: customer_name required."); continue
                    items_str = row["items_json"]
                    try:
                        items = json.loads(items_str)
                        if not isinstance(items, list) or not items:
                            raise ValueError("items_json must be non-empty array")
                    except Exception as e:
                        errors.append(f"Row {i}: invalid items_json ({e})."); continue
                    date_str = row["order_date"] or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    status = row["status"] or "Pending"
                    if status not in ORDER_STATUSES: status = "Pending"
                    order_items = []
                    for j, it in enumerate(items, 1):
                        sku = (it.get("product_sku") or "").strip()
                        qty = int(it.get("quantity", 0) or 0)
                        price_at = float(it.get("price_at_order", 0) or 0)
                        if not sku or qty <= 0 or price_at <= 0:
                            errors.append(f"Row {i} item {j}: invalid item (sku/qty/price)."); order_items = []; break
                        p = con.execute("SELECT id, quantity FROM products WHERE sku=?", (sku,)).fetchone()
                        if not p:
                            errors.append(f"Row {i} item {j}: unknown SKU '{sku}'."); order_items = []; break
                        pid, avail = p
                        if avail < qty:
                            errors.append(f"Row {i} item {j}: not enough stock for {sku} (avail {avail})."); order_items = []; break
                        order_items.append({"product_id": pid, "sku": sku, "quantity": qty, "price_at_order": price_at})
                    if not order_items: continue
                    cur = con.execute("INSERT INTO orders(customer_name, order_date, status) VALUES (?,?,?)",
                                      (cust, date_str, status))
                    oid = cur.lastrowid
                    for it in order_items:
                        con.execute("""INSERT INTO order_items(order_id, product_id, quantity, price_at_order)
                                       VALUES (?,?,?,?)""", (oid, it["product_id"], it["quantity"], it["price_at_order"]))
                        adjust_stock(con, it["product_id"], -it["quantity"])
                    created += 1
            msg = f"Orders import complete.\nCreated: {created}"
            if errors:
                msg += f"\nErrors: {len(errors)} (up to 20):\n" + "\n".join(errors[:20])
                messagebox.showwarning("Import completed with issues", msg)
            else:
                messagebox.showinfo("Import", msg)
            self._refresh_all()
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
    def _build_settings(self):
        f = self.tab_settings
        f.columnconfigure(0, weight=1)
        lbl = ttk.Label(f, text="Validation & Update Policies", font=("Segoe UI", 12, "bold"))
        lbl.grid(row=0, column=0, sticky="w", padx=8, pady=(8,4))
        box = ttk.LabelFrame(f, text="Validation", padding=8)
        box.grid(row=1, column=0, sticky="ew", padx=8, pady=6)
        self.var_supplier_name_unique = tk.BooleanVar(value=settings_get_bool("supplier_name_unique"))
        self.var_require_supplier_id = tk.BooleanVar(value=settings_get_bool("require_supplier_id_on_product"))
        self.var_enforce_price_gt0 = tk.BooleanVar(value=settings_get_bool("enforce_price_gt_zero_on_import"))
        self.var_enforce_qty_ge0 = tk.BooleanVar(value=settings_get_bool("enforce_quantity_ge_zero_on_import"))
        ttk.Checkbutton(box, text="Supplier name must be unique (UI policy)", variable=self.var_supplier_name_unique).grid(sticky="w", padx=4, pady=2)
        ttk.Checkbutton(box, text="Require supplier on product", variable=self.var_require_supplier_id).grid(sticky="w", padx=4, pady=2)
        ttk.Checkbutton(box, text="Import requires price > 0", variable=self.var_enforce_price_gt0).grid(sticky="w", padx=4, pady=2)
        ttk.Checkbutton(box, text="Import requires quantity ≥ 0", variable=self.var_enforce_qty_ge0).grid(sticky="w", padx=4, pady=2)
        box2 = ttk.LabelFrame(f, text="Product update policy on import", padding=8)
        box2.grid(row=2, column=0, sticky="ew", padx=8, pady=6)
        self.var_allow_update_supplier = tk.BooleanVar(value=settings_get_bool("allow_update_product_supplier"))
        self.var_allow_update_price = tk.BooleanVar(value=settings_get_bool("allow_update_product_price"))
        self.var_allow_update_name = tk.BooleanVar(value=settings_get_bool("allow_update_product_name"))
        ttk.Checkbutton(box2, text="Allow supplier update", variable=self.var_allow_update_supplier).grid(sticky="w", padx=4, pady=2)
        ttk.Checkbutton(box2, text="Allow price update", variable=self.var_allow_update_price).grid(sticky="w", padx=4, pady=2)
        ttk.Checkbutton(box2, text="Allow name update", variable=self.var_allow_update_name).grid(sticky="w", padx=4, pady=2)
        btns = ttk.Frame(f); btns.grid(row=3, column=0, sticky="e", padx=8, pady=8)
        ttk.Button(btns, text="Save Settings", command=self._save_settings, style="Accent.TButton").pack(side="right")
        note = ttk.Label(f, text="Freight rates & default percent can be edited under Supplier ➜ double-click ➜ 'Freight' tab.",
                         foreground="#666")
        note.grid(row=4, column=0, sticky="w", padx=8, pady=(0,8))
    def _save_settings(self):
        try:
            settings_set_bool("supplier_name_unique", bool(self.var_supplier_name_unique.get()))
            settings_set_bool("require_supplier_id_on_product", bool(self.var_require_supplier_id.get()))
            settings_set_bool("enforce_price_gt_zero_on_import", bool(self.var_enforce_price_gt0.get()))
            settings_set_bool("enforce_quantity_ge_zero_on_import", bool(self.var_enforce_qty_ge0.get()))
            settings_set_bool("allow_update_product_supplier", bool(self.var_allow_update_supplier.get()))
            settings_set_bool("allow_update_product_price", bool(self.var_allow_update_price.get()))
            settings_set_bool("allow_update_product_name", bool(self.var_allow_update_name.get()))
            messagebox.showinfo("Settings", "Settings saved.")
        except Exception as e:
            messagebox.showerror("Settings", f"Failed to save settings:\n{e}")
    def _save_template_file(self, title: str, file_cols: list[str]):
        path = filedialog.asksaveasfilename(title=title,
                                            defaultextension=".csv",
                                            filetypes=[("CSV","*.csv")] + ([("Excel","*.xlsx")] if PANDAS_OK else []))
        if not path: return
        try:
            rows = [{c: "" for c in file_cols}]
            write_table_to_file(path, rows, file_cols)
            messagebox.showinfo("Template", f"Template written to:\n{path}")
        except Exception as e:
            messagebox.showerror("Template failed", str(e))
    def _export_data_file(self, title: str, file_cols: list[str], data_rows: list[dict]):
        path = filedialog.asksaveasfilename(title=title,
                                            defaultextension=".csv",
                                            filetypes=[("CSV","*.csv")] + ([("Excel","*.xlsx")] if PANDAS_OK else []))
        if not path: return
        try:
            write_table_to_file(path, data_rows, file_cols)
            messagebox.showinfo("Export", f"Data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))
    def _refresh_all(self):
        self._refresh_dashboard()
        self._refresh_orders()
        self._refresh_suppliers()
        self._refresh_products()
        self._clear_search_products()
if __name__ == "__main__":
    db_init()
    App().mainloop()

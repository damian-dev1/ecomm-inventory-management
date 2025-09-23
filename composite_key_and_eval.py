#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import json
import re
import sys
from math import floor, ceil
from pathlib import Path
from typing import List, Dict, Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import datetime as dt
from openpyxl import Workbook

import pandas as pd
from sqlalchemy import create_engine, text

# =========================
# THEME
# =========================

class ThemeManager:
    def __init__(self, root: tk.Tk, *, load_tcl: bool = False, tcl_path: str | None = None):
        self.root = root
        self.colors = {
            "bg": "#1a1b26",
            "fg": "#c0caf5",
            "muted": "#a9b1d6",
            "accent": "#7aa2f7",
            "button": "#2f3549",
            "panel": "#1b1d2b",
            "border": "#2a2f37",
            "success": "#9ece6a",
            "warn": "#e0af68",
            "error": "#f7768e",
            "info": "#7aa2f7",
            "field": "#222436",
            "active": "#3b4261",
            "tree_bg": "#1a1b26",
            "tree_alt": "#171822",
            "tree_sel": "#283457",
            "hd_bg": "#1e2130",
            "hd_fg": "#93a3e7",
        }
        if load_tcl and tcl_path:
            try:
                self.root.tk.call("source", tcl_path)
                self.root.tk.call("set_theme")
            except tk.TclError:
                pass
        self._apply_base()
        self._apply_combobox_popdown_dark()

    def _apply_base(self):
        c = self.colors
        self.root.configure(bg=c["bg"])
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=c["bg"], foreground=c["fg"])
        style.configure("TFrame", background=c["bg"])
        style.configure("TLabelframe", background=c["bg"], foreground=c["fg"], bordercolor=c["border"])
        style.configure("TLabelframe.Label", background=c["bg"], foreground=c["fg"])
        style.configure("TLabel", background=c["bg"], foreground=c["fg"])
        style.configure("TSeparator", background=c["border"])
        style.configure("TButton", background=c["button"], foreground=c["fg"], padding=6)
        style.map("TButton", background=[("active", c["active"])])
        style.configure("Accent.TButton", background=c["accent"], foreground="#0e101a")
        style.map("Accent.TButton", background=[("active", "#5e81ac")])
        style.configure("TCheckbutton", background=c["bg"], foreground=c["fg"])
        style.configure("TRadiobutton", background=c["bg"], foreground=c["fg"])
        style.configure("TEntry", fieldbackground=c["field"], foreground=c["fg"], insertcolor=c["accent"])
        style.configure("TSpinbox", fieldbackground=c["field"], foreground=c["fg"])
        style.configure(
            "TCombobox",
            fieldbackground=c["field"],
            foreground=c["fg"],
            selectbackground=c["accent"],
            selectforeground="#10121a",
            arrowsize=14,
            bordercolor=c["border"],
            lightcolor=c["border"],
            darkcolor=c["border"]
        )
        style.map(
            "TCombobox",
            fieldbackground=[
                ("readonly", c["field"]),
                ("!readonly", c["field"]),
                ("focus", c["field"]),
                ("active", c["field"])
            ],
            foreground=[
                ("readonly", c["fg"]),
                ("focus", c["fg"]),
                ("active", c["fg"])
            ],
            background=[
                ("readonly", c["bg"]),
                ("focus", c["bg"]),
                ("active", c["bg"])
            ],
            arrowcolor=[
                ("readonly", c["fg"]),
                ("focus", c["fg"]),
                ("active", c["fg"])
            ],
            bordercolor=[
                ("focus", c["active"]),
                ("!focus", c["border"])
            ]
        )
        style.configure(
            "Treeview",
            background=c["tree_bg"],
            fieldbackground=c["tree_bg"],
            foreground=c["fg"],
            rowheight=22,
            bordercolor=c["border"],
            lightcolor=c["border"],
            darkcolor=c["border"],
        )
        style.map(
            "Treeview",
            background=[("selected", c["tree_sel"])],
            foreground=[("selected", c["fg"])],
        )
        style.configure("Treeview.Heading", background=c["hd_bg"], foreground=c["hd_fg"], relief="flat", padding=4)
        style.map("Treeview.Heading", background=[("active", c["active"])])
        self.style = style

    def _apply_combobox_popdown_dark(self):
        c = self.colors
        self.root.option_add('*TCombobox*Listbox.background', c["field"])
        self.root.option_add('*TCombobox*Listbox.foreground', c["fg"])
        self.root.option_add('*TCombobox*Listbox.selectBackground', c["tree_sel"])
        self.root.option_add('*TCombobox*Listbox.selectForeground', c["fg"])
        self.root.option_add('*TCombobox*Listbox.borderWidth', 0)
        self.root.option_add('*TCombobox*Listbox.relief', 'flat')
        self.root.option_add('*ComboboxPopdownFrame*background', c["bg"])

    def apply_to_toplevel(self, top: tk.Toplevel):
        top.configure(bg=self.colors["bg"])


# =========================
# CONFIG / CONSTANTS
# =========================

CONFIG_FILE = Path("tolerance_config.json")
COMPARISON_KEY = "Stock Qty Comparison"

ALL_COLUMNS = [
    "Timestamp", "SKU", "Ecommerce Qty", "Warehouse Qty", "Use Low Stock Logic",
    "Low Stock Max Qty", "Applied Severity", "Tolerance (%)", "Tolerance Base",
    "Tolerance Rounding", "Tolerance Units", "Delta", "Ecommerce Status",
    "Warehouse Status", "Availability", COMPARISON_KEY, "Reason"
]

DEFAULT_VISIBLE_COLUMNS = [
    "Timestamp", "SKU", "Ecommerce Qty", "Warehouse Qty", "Delta", "Applied Severity", COMPARISON_KEY, "Reason"
]

DEFAULT_CONFIG = {
    "rounding": "floor",
    "base_method": "max",
    "severity_order": ["Critical", "High", "Medium", "Low", "Very Low"],
    "severity_map": {
        "Critical":  {"pct": 0.00, "min_units": 0, "max_units": 0},
        "High":      {"pct": 0.00, "min_units": 0, "max_units": 0},
        "Medium":    {"pct": 0.05, "min_units": 1, "max_units": None},
        "Low":       {"pct": 0.10, "min_units": 1, "max_units": None},
        "Very Low":  {"pct": 0.20, "min_units": 1, "max_units": None}
    },
    "auto_equal_status_severity": "Medium",
    "visible_columns": DEFAULT_VISIBLE_COLUMNS,
    "qty_compare_mode": "strict"
}

def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            merged = DEFAULT_CONFIG.copy()
            merged.update(cfg)
            merged["severity_map"] = {**DEFAULT_CONFIG["severity_map"], **cfg.get("severity_map", {})}
            if "severity_order" not in merged:
                merged["severity_order"] = DEFAULT_CONFIG["severity_order"]
            vis = merged.get("visible_columns") or DEFAULT_VISIBLE_COLUMNS
            merged["visible_columns"] = [c for c in vis if c in ALL_COLUMNS] or DEFAULT_VISIBLE_COLUMNS
            return merged
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    to_save = dict(cfg)
    to_save["visible_columns"] = [c for c in cfg.get("visible_columns", DEFAULT_VISIBLE_COLUMNS) if c in ALL_COLUMNS]
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(to_save, f, indent=2)


# =========================
# UI HELPERS
# =========================

class ContextMenuManager:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.menu = tk.Menu(root, tearoff=False)
        self.menu.add_command(label="Cut", command=self._cut)
        self.menu.add_command(label="Copy", command=self._copy)
        self.menu.add_command(label="Paste", command=self._paste)
        self.menu.add_separator()
        self.menu.add_command(label="Select All", command=self._select_all)
        self.target = None
        self.root.bind_all("<Button-3>", self._show, add="+")
        self.root.bind_all("<Control-Button-1>", self._show, add="+")

    def _show(self, event):
        widget = event.widget
        if isinstance(widget, (tk.Entry, ttk.Entry, tk.Text)):
            self.target = widget
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.menu.grab_release()

    def _cut(self): self._event("<<Cut>>")
    def _copy(self): self._event("<<Copy>>")
    def _paste(self): self._event("<<Paste>>")

    def _select_all(self):
        w = self.target
        if isinstance(w, tk.Text):
            w.tag_add("sel", "1.0", "end-1c")
        else:
            w.selection_range(0, "end")

    def _event(self, name):
        if self.target:
            self.target.event_generate(name)


class BaseDialog(tk.Toplevel):
    def __init__(self, parent, theme: ThemeManager, title: str):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        theme.apply_to_toplevel(self)
        self.body = ttk.Frame(self, padding=10)
        self.body.pack(fill="both", expand=True)


class SettingsDialog(BaseDialog):
    def __init__(self, parent, theme: ThemeManager, cfg, on_save):
        super().__init__(parent, theme, "Settings")
        self.cfg = json.loads(json.dumps(cfg))
        self.on_save = on_save
        pad = {"padx": 6, "pady": 4}
        row = 0

        ttk.Label(self.body, text="Qty Compare Mode").grid(row=row, column=0, **pad, sticky="e")
        self.mode_var = tk.StringVar(value=self.cfg.get("qty_compare_mode", "strict"))
        ttk.Combobox(self.body, textvariable=self.mode_var, state="readonly",
                     values=["strict", "tolerant"], width=12).grid(row=row, column=1, **pad, sticky="w")

        ttk.Label(self.body, text="Base Method").grid(row=row, column=2, **pad, sticky="e")
        self.base_var = tk.StringVar(value=self.cfg.get("base_method", "max"))
        ttk.Combobox(self.body, textvariable=self.base_var, state="readonly",
                     values=["max", "avg", "min", "ecommerce", "warehouse"], width=14).grid(row=row, column=3, **pad, sticky="w")
        row += 1

        ttk.Label(self.body, text="Rounding").grid(row=row, column=0, **pad, sticky="e")
        self.round_var = tk.StringVar(value=self.cfg.get("rounding", "floor"))
        ttk.Combobox(self.body, textvariable=self.round_var, state="readonly",
                     values=["floor", "ceil", "round"], width=10).grid(row=row, column=1, **pad, sticky="w")

        ttk.Label(self.body, text="Auto (equal statuses)").grid(row=row, column=2, **pad, sticky="e")
        self.auto_equal_var = tk.StringVar(value=self.cfg.get("auto_equal_status_severity", "Medium"))
        ttk.Combobox(self.body, textvariable=self.auto_equal_var, state="readonly",
                     values=DEFAULT_CONFIG["severity_order"], width=14).grid(row=row, column=3, **pad, sticky="w")
        row += 1

        ttk.Separator(self.body, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="ew", pady=6)
        row += 1

        btns = ttk.Frame(self.body)
        btns.grid(row=row, column=0, columnspan=4, pady=(8, 0), sticky="e")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Save", style="Accent.TButton", command=self._save).pack(side="right")

    def _save(self):
        try:
            self.cfg["qty_compare_mode"] = self.mode_var.get()
            self.cfg["base_method"] = self.base_var.get()
            self.cfg["rounding"] = self.round_var.get()
            self.cfg["auto_equal_status_severity"] = self.auto_equal_var.get()
            save_config(self.cfg)
            self.on_save(self.cfg)
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Input", "Ensure numeric fields are valid.", parent=self)


# =========================
# CHECKER (Single SKU)
# =========================

class InventoryChecker:
    def __init__(self, sku, ecommerce_qty, warehouse_qty, use_low_stock, low_stock_max, config):
        self.sku = str(sku).strip()
        self.ec = int(ecommerce_qty)
        self.wh = int(warehouse_qty)
        self.use_low_stock = bool(use_low_stock)
        self.low_stock_max = int(low_stock_max)
        self.cfg = config

    def get_status(self, qty: int) -> str:
        if qty == 0:
            return "Out of Stock"
        if self.use_low_stock and 1 <= qty <= self.low_stock_max:
            return "Low Stock"
        return "In Stock"

    def get_availability(self, ec_status: str, wh_status: str) -> str:
        if ec_status != "Out of Stock" and wh_status == "Out of Stock":
            return "Available in Ecommerce only"
        if ec_status == "Out of Stock" and wh_status != "Out of Stock":
            return "Available in Warehouse only"
        if ec_status != "Out of Stock" and wh_status != "Out of Stock":
            return "Available in both Ecommerce and Warehouse"
        return "Out of Stock everywhere"

    def decide_severity_auto(self, ec_status: str, wh_status: str) -> str:
        if (ec_status == "Out of Stock") != (wh_status == "Out of Stock"):
            return "Critical"
        if ec_status != wh_status:
            return "High"
        return self.cfg.get("auto_equal_status_severity", "Medium")

    def tolerance_units(self, severity: str):
        m = DEFAULT_CONFIG["severity_map"][severity]
        pct = float(m.get("pct", 0.0))
        base_method = self.cfg.get("base_method", "max")
        rounding_mode = self.cfg.get("rounding", "floor")
        if base_method == "max":
            base = max(self.ec, self.wh)
        elif base_method == "min":
            base = min(self.ec, self.wh)
        elif base_method == "avg":
            base = (self.ec + self.wh) / 2.0
        elif base_method == "ecommerce":
            base = float(self.ec)
        elif base_method == "warehouse":
            base = float(self.wh)
        else:
            base = max(self.ec, self.wh)
        raw = base * pct
        if rounding_mode == "ceil":
            units = ceil(raw)
        elif rounding_mode == "round":
            units = round(raw)
        else:
            units = floor(raw)
        min_units = int(m.get("min_units", 0) or 0)
        max_units = m.get("max_units", None)
        if max_units is not None:
            max_units = int(max_units)
        units = max(min_units, units)
        if max_units is not None:
            units = min(units, max_units)
        return units, pct, base_method, rounding_mode

    def evaluate(self):
        ec_status = self.get_status(self.ec)
        wh_status = self.get_status(self.wh)
        applied_sev = self.decide_severity_auto(ec_status, wh_status)
        tol_units, tol_pct, base_method, rounding_mode = self.tolerance_units(applied_sev)
        delta = abs(self.ec - self.wh)
        mode = self.cfg.get("qty_compare_mode", "strict")
        if mode == "strict":
            if self.ec == self.wh:
                label = "Match"
                reason = "Quantities are equal."
            else:
                label = "Mismatch"
                reason = f"Quantities differ (Δ={delta})."
        else:
            if ec_status != wh_status and tol_units == 0:
                label = "Mismatch"
                reason = "Statuses diverge; tolerance is 0 for this severity."
            else:
                if delta <= tol_units:
                    label = "Match"
                    reason = f"Delta ≤ tolerance ({delta} ≤ {tol_units})."
                else:
                    label = "Mismatch"
                    reason = f"Delta exceeds tolerance ({delta} > {tol_units})."
        return {
            "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "SKU": self.sku,
            "Ecommerce Qty": self.ec,
            "Warehouse Qty": self.wh,
            "Use Low Stock Logic": self.use_low_stock,
            "Low Stock Max Qty": self.low_stock_max,
            "Applied Severity": applied_sev,
            "Tolerance (%)": tol_pct,
            "Tolerance Base": base_method,
            "Tolerance Rounding": rounding_mode,
            "Tolerance Units": tol_units,
            "Delta": delta,
            "Ecommerce Status": ec_status,
            "Warehouse Status": wh_status,
            "Availability": self.get_availability(ec_status, wh_status),
            COMPARISON_KEY: label,
            "Reason": reason,
        }


# =========================
# CHECKER TAB
# =========================

class CheckerTab(ttk.Frame):
    def __init__(self, parent, theme: ThemeManager, cfg: dict):
        super().__init__(parent, padding=(10, 8))
        self.theme = theme
        self.cfg = cfg
        self.results: list[dict] = []

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Evaluate", style="Accent.TButton", command=self.evaluate_one).pack(side="left")
        ttk.Button(toolbar, text="Settings…", command=self.open_settings).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Choose Columns…", command=self.choose_columns).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Export Excel", command=self.export_excel).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Clear Results", command=self.clear_results).pack(side="left", padx=(6, 0))

        inputs = ttk.Frame(self)
        inputs.pack(fill="x", pady=(0, 6))
        self.sku_entry = self._labeled_entry(inputs, "SKU", "TEST-SKU-001")
        self.ec_entry  = self._labeled_entry(inputs, "EC Qty", "100")
        self.wh_entry  = self._labeled_entry(inputs, "WH Qty", "102")
        self.use_low_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(inputs, text="Low Stock", variable=self.use_low_var, command=self._toggle_low_stock).pack(side="left", padx=(8, 0))
        ttk.Label(inputs, text="Max").pack(side="left", padx=(6, 2))
        self.low_max_var = tk.IntVar(value=5)
        self.low_spin = ttk.Spinbox(inputs, from_=1, to=9999, width=6, textvariable=self.low_max_var, state="readonly")
        self.low_spin.pack(side="left")

        table_box = ttk.Frame(self)
        table_box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table_box, columns=self.cfg["visible_columns"], show="headings", height=14)
        self._apply_tree_columns(self.cfg["visible_columns"])
        yscroll = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_box.grid_columnconfigure(0, weight=1)
        table_box.grid_rowconfigure(0, weight=1)
        self.tree.tag_configure("oddrow", background=self.theme.colors["tree_alt"])
        self.tree.tag_configure("evenrow", background=self.theme.colors["tree_bg"])
        self.tree.tag_configure("mismatch", foreground=self.theme.colors["error"])
        self.tree.tag_configure("match", foreground=self.theme.colors["success"])

        self.status = ttk.Label(self, text="Ready.")
        self.status.pack(fill="x", pady=(6, 0))

        self.bind_all("<Return>", lambda e: self.evaluate_one())
        self._toggle_low_stock()

    def _labeled_entry(self, parent, label, initial="", width=14):
        frm = ttk.Frame(parent)
        frm.pack(side="left", padx=(0, 8))
        ttk.Label(frm, text=label).pack(side="left", padx=(0, 4))
        ent = ttk.Entry(frm, width=width)
        ent.pack(side="left")
        if initial:
            ent.insert(0, initial)
        return ent

    def _toggle_low_stock(self):
        self.low_spin.configure(state="readonly" if self.use_low_var.get() else "disabled")

    def _get_int(self, entry: ttk.Entry, name: str) -> int:
        val = entry.get().strip()
        try:
            return int(val)
        except ValueError:
            messagebox.showerror("Invalid Input", f"{name} must be an integer.")
            entry.focus_set()
            raise

    def _apply_tree_columns(self, visible_cols: list[str]):
        self.tree["columns"] = visible_cols
        for col in visible_cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=max(90, len(col) * 7 + 10), anchor="w", stretch=True)

    def _refresh_tree_rows(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for idx, r in enumerate(self.results):
            self._insert_row(idx, r)

    def evaluate_one(self):
        try:
            sku = (self.sku_entry.get().strip() or "UNKNOWN-SKU")
            ec = self._get_int(self.ec_entry, "EC Qty")
            wh = self._get_int(self.wh_entry, "WH Qty")
            use_low = self.use_low_var.get()
            low_max = self.low_max_var.get() if use_low else 0
            checker = InventoryChecker(sku, ec, wh, use_low, low_max, self.cfg)
            result = checker.evaluate()
            self.results.append(result)
            self._insert_row(len(self.results)-1, result)
            self.status.configure(text=f"{sku}: {result[COMPARISON_KEY]} • Δ={result['Delta']} • Tol={result['Tolerance Units']}")
        except Exception:
            return

    def clear_results(self):
        if messagebox.askyesno("Clear Results", "Clear all results?"):
            self.results.clear()
            self._refresh_tree_rows()
            self.status.configure(text="Results cleared.")

    def _insert_row(self, idx: int, result: dict):
        values = [result.get(col, "") for col in self.tree["columns"]]
        tags = ["oddrow" if idx % 2 else "evenrow"]
        tags.append("mismatch" if result.get(COMPARISON_KEY) == "Mismatch" else "match")
        self.tree.insert("", "end", values=values, tags=tuple(tags))

    def open_settings(self):
        def on_save(cfg):
            self.cfg = cfg
            self.status.configure(text="Settings saved.")
        SettingsDialog(self, self.theme, self.cfg, on_save)

    def choose_columns(self):
        def on_save(selected_cols: list[str]):
            self.cfg["visible_columns"] = selected_cols
            save_config(self.cfg)
            self._apply_tree_columns(selected_cols)
            self._refresh_tree_rows()
            self.status.configure(text=f"Columns updated ({len(selected_cols)} shown).")
        ColumnChooserDialog(self, self.theme, self.cfg.get("visible_columns", DEFAULT_VISIBLE_COLUMNS), on_save)

    def export_excel(self):
        if not self.results:
            messagebox.showinfo("Export", "No results to export yet.")
            return
        default_name = f"inventory_results_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Export to Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Results"
            ws.append(ALL_COLUMNS)
            for r in self.results:
                ws.append([r.get(c, "") for c in ALL_COLUMNS])
            wb.save(path)
            messagebox.showinfo("Export", f"Exported {len(self.results)} rows to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))


class ColumnChooserDialog(BaseDialog):
    def __init__(self, parent, theme: ThemeManager, current_visible: list[str], on_save):
        super().__init__(parent, theme, "Choose Columns")
        self.on_save = on_save
        pad = {"padx": 4, "pady": 2}
        self.vars = {}

        ttk.Label(self.body, text="Select columns to display:").grid(row=0, column=0, sticky="w", **pad)
        grid = ttk.Frame(self.body)
        grid.grid(row=1, column=0, sticky="nsew", **pad)

        cols = list(ALL_COLUMNS)
        per_col = (len(cols) + 2) // 3
        for i, col in enumerate(cols):
            r = i % per_col
            c = i // per_col
            var = tk.BooleanVar(value=(col in current_visible))
            self.vars[col] = var
            ttk.Checkbutton(grid, text=col, variable=var).grid(row=r, column=c, sticky="w", padx=6, pady=2)

        presets = ttk.Frame(self.body)
        presets.grid(row=2, column=0, sticky="w", **pad)
        ttk.Button(presets, text="Main Stats", command=self._preset_main).pack(side="left", padx=2)
        ttk.Button(presets, text="All", command=self._preset_all).pack(side="left", padx=2)
        ttk.Button(presets, text="None", command=self._preset_none).pack(side="left", padx=2)

        btns = ttk.Frame(self.body)
        btns.grid(row=3, column=0, sticky="e", pady=(8,0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Apply", style="Accent.TButton", command=self._apply).pack(side="right")

    def _preset_main(self):
        for col, var in self.vars.items():
            var.set(col in DEFAULT_VISIBLE_COLUMNS)

    def _preset_all(self):
        for var in self.vars.values():
            var.set(True)

    def _preset_none(self):
        for var in self.vars.values():
            var.set(False)

    def _apply(self):
        selected = [c for c, v in self.vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Columns", "Select at least one column.")
            return
        self.on_save(selected)
        self.destroy()


# =========================
# COMPOSITE-KEY COMPARE TAB
# =========================

PROFILES_FILE = "ck_profiles.json"

def sniff_delimiter(path: Path) -> str:
    with path.open("r", encoding="utf-8", errors="replace") as f:
        sample = f.read(4096)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        return ","

def normalize_piece(s: str, *, case: str, trim: bool, collapse_spaces: bool, zero_pad: int) -> str:
    if pd.isna(s):
        s = ""
    s = str(s)
    if trim:
        s = s.strip()
    if collapse_spaces:
        s = re.sub(r"\s+", " ", s)
    if case == "lower":
        s = s.casefold()
    elif case == "upper":
        s = s.upper()
    if zero_pad and s.isdigit():
        s = s.zfill(zero_pad)
    return s

def build_composite_key_series(df: pd.DataFrame, cols: List[str], norm: Dict) -> pd.Series:
    use = [c for c in cols if c and c in df.columns]
    if not use:
        return pd.Series([""], index=df.index)
    parts = []
    for c in use:
        parts.append(df[c].astype(str).map(lambda x: normalize_piece(
            x,
            case=norm["case"],
            trim=norm["trim"],
            collapse_spaces=norm["collapse_spaces"],
            zero_pad=norm["zero_pad"],
        )))
    out = parts[0]
    for p in parts[1:]:
        out = out + "|" + p
    return out

def coerce_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0).astype(float)

class ChipBar(ttk.Frame):
    def __init__(self, parent, on_change, max_chips=3):
        super().__init__(parent)
        self.columns: List[str] = []
        self.on_change = on_change
        self.max_chips = max_chips

    def set_values(self, values: List[str]):
        self.columns = [v for v in values if v]
        self._render()
        self.on_change(self.columns)

    def get_values(self) -> List[str]:
        return list(self.columns)

    def add(self, col: str):
        if not col:
            return
        if col in self.columns:
            return
        if len(self.columns) >= self.max_chips:
            messagebox.showwarning("Key Builder", f"Max {self.max_chips} columns for composite key.")
            return
        self.columns.append(col)
        self._render()
        self.on_change(self.columns)

    def remove(self, col: str):
        self.columns = [c for c in self.columns if c != col]
        self._render()
        self.on_change(self.columns)

    def _render(self):
        for w in self.winfo_children():
            w.destroy()
        for col in self.columns:
            chip = ttk.Frame(self, padding=(6, 2))
            ttk.Label(chip, text=col).pack(side="left")
            ttk.Button(chip, text="×", width=2, command=lambda c=col: self.remove(c)).pack(side="left", padx=(6, 0))
            chip.pack(side="left", padx=4)

class SourceCard(ttk.Labelframe):
    def __init__(self, parent, title: str, theme: ThemeManager):
        super().__init__(parent, text=title, padding=10)
        self.df: Optional[pd.DataFrame] = None
        self.title = title
        self.theme = theme

        self.btn = ttk.Button(self, text="Load CSV", command=self._on_load)
        self.btn.grid(row=0, column=0, sticky="w")

        self.badge = ttk.Label(self, text="No file", foreground=self.theme.colors["accent"])
        self.badge.grid(row=0, column=1, sticky="w", padx=(10, 0))

        self.col_var = tk.StringVar()
        self.col_dd = ttk.Combobox(self, textvariable=self.col_var, state="readonly", width=32)
        self.col_dd.grid(row=1, column=0, sticky="w", pady=(8, 2))
        self.add_btn = ttk.Button(self, text="Add →", command=self._add_selected, width=8)
        self.add_btn.grid(row=1, column=1, sticky="w", padx=(8, 0))

        ttk.Label(self, text="Composite Key:").grid(row=2, column=0, sticky="w", pady=(10, 2))
        self.chips = ChipBar(self, on_change=lambda cols: None)
        self.chips.grid(row=3, column=0, columnspan=2, sticky="w")

        ttk.Label(self, text="Quantity Column:").grid(row=4, column=0, sticky="w", pady=(10, 2))
        self.qty_var = tk.StringVar()
        self.qty_dd = ttk.Combobox(self, textvariable=self.qty_var, state="readonly", width=32)
        self.qty_dd.grid(row=5, column=0, sticky="w")

        norm_fr = ttk.Frame(self, padding=(0, 8, 0, 0))
        norm_fr.grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Label(norm_fr, text="Normalization:").grid(row=0, column=0, sticky="w", pady=(6, 2))

        self.case_var = tk.StringVar(value="lower")
        ttk.Label(norm_fr, text="Case").grid(row=1, column=0, sticky="w")
        ttk.Combobox(norm_fr, textvariable=self.case_var, state="readonly",
                     values=["lower", "upper", "as-is"], width=8).grid(row=1, column=1, sticky="w", padx=(6, 12))
        self.trim_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(norm_fr, text="Trim", variable=self.trim_var).grid(row=1, column=2, sticky="w")
        self.collapse_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(norm_fr, text="Collapse spaces", variable=self.collapse_var).grid(row=1, column=3, sticky="w", padx=(8, 0))

        ttk.Label(norm_fr, text="Zero-pad").grid(row=1, column=4, sticky="w", padx=(12, 0))
        self.pad_var = tk.IntVar(value=0)
        ttk.Spinbox(norm_fr, from_=0, to=12, textvariable=self.pad_var, width=4).grid(row=1, column=5, sticky="w", padx=(6, 0))

        ttk.Label(self, text="Key Preview (first 10):").grid(row=7, column=0, sticky="w", pady=(10, 2))
        self.preview = tk.Text(self, height=6, width=44, bg=self.theme.colors["tree_bg"], fg=self.theme.colors["fg"], relief="flat")
        self.preview.grid(row=8, column=0, columnspan=2, sticky="we")
        self.grid_columnconfigure(0, weight=1)

    def _add_selected(self):
        col = self.col_var.get()
        if not col:
            return
        self.chips.add(col)
        self._refresh_preview()

    def _on_load(self):
        fpath = filedialog.asksaveasfilename if False else filedialog.askopenfilename
        path = fpath(title=f"Select {self.title} CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        p = Path(path)
        delim = sniff_delimiter(p)
        try:
            df = pd.read_csv(p, dtype=str, sep=delim, low_memory=False)
        except Exception as e:
            messagebox.showerror("Load CSV", f"Failed to load file:\n{e}")
            return
        if df.empty or not list(df.columns):
            messagebox.showwarning("Load CSV", "No data/columns detected.")
            return
        self.df = df
        cols = list(df.columns)
        self.col_dd["values"] = cols
        self.qty_dd["values"] = cols
        default_ids = [c for c in cols if c.lower() in {"sku", "id", "product_id"} or "sku" in c.lower()]
        default_qty = [c for c in cols if "qty" in c.lower() or "quantity" in c.lower() or c.lower() in {"stock", "soh"}]
        if default_ids:
            self.chips.set_values([default_ids[0]])
        else:
            self.chips.set_values([cols[0]])
        self.qty_var.set(default_qty[0] if default_qty else (cols[1] if len(cols) > 1 else cols[0]))
        self.badge.configure(text=f"Loaded ({len(df):,} rows, {len(cols)} cols)")
        self._refresh_preview()

    def get_norm(self) -> Dict:
        return {
            "case": self.case_var.get(),
            "trim": bool(self.trim_var.get()),
            "collapse_spaces": bool(self.collapse_var.get()),
            "zero_pad": int(self.pad_var.get() or 0),
        }

    def get_key_cols(self) -> List[str]:
        return self.chips.get_values()

    def get_qty_col(self) -> str:
        return self.qty_var.get().strip()

    def _refresh_preview(self):
        self.preview.delete("1.0", "end")
        if self.df is None:
            return
        keys = build_composite_key_series(self.df.head(10), self.get_key_cols(), self.get_norm())
        for k in keys.tolist():
            self.preview.insert("end", f"{k}\n")


def load_profiles() -> Dict:
    p = Path(PROFILES_FILE)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_profiles(data: Dict):
    Path(PROFILES_FILE).write_text(json.dumps(data, indent=2), encoding="utf-8")


class CompositeResultsPane(ttk.Frame):
    def __init__(self, parent, theme: ThemeManager):
        super().__init__(parent, padding=10)
        self.theme = theme

        tbar = ttk.Frame(self)
        tbar.grid(row=0, column=0, sticky="ew")
        for i in range(8):
            tbar.grid_columnconfigure(i, weight=1)

        self.only_diff = tk.BooleanVar(value=False)
        ttk.Checkbutton(tbar, text="Only differences", variable=self.only_diff).grid(row=0, column=0, sticky="w")

        ttk.Label(tbar, text="Presence").grid(row=0, column=1, sticky="e")
        self.presence = tk.StringVar(value="All")
        ttk.Combobox(tbar, textvariable=self.presence, state="readonly",
                     values=["All", "Both", "Only in Warehouse", "Only in Ecommerce"], width=18)\
            .grid(row=0, column=2, sticky="w", padx=(6, 12))

        ttk.Label(tbar, text="Diff").grid(row=0, column=3, sticky="e")
        self.diff_op = tk.StringVar(value="!=")
        ttk.Combobox(tbar, textvariable=self.diff_op, state="readonly",
                     values=["!=", ">", "<", ">=", "<=", "=="], width=4)\
            .grid(row=0, column=4, sticky="w", padx=(6, 4))
        self.diff_thr = tk.DoubleVar(value=0.0)
        ttk.Entry(tbar, textvariable=self.diff_thr, width=8).grid(row=0, column=5, sticky="w")

        self.btn_compare = ttk.Button(tbar, text="Compare", width=14)
        self.btn_export = ttk.Button(tbar, text="Export CSV", width=14)
        self.btn_compare.grid(row=0, column=6, sticky="e", padx=(12, 6))
        self.btn_export.grid(row=0, column=7, sticky="e", padx=(6, 0))

        results = ttk.Frame(self, padding=(0, 8, 0, 0))
        results.grid(row=1, column=0, sticky="nsew")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        cols = ("Key", "Warehouse Qty", "Ecommerce Qty", "Difference", "Presence")
        self.tree = ttk.Treeview(results, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c, command=lambda col=c: self._sort_tree_by(col, False))
            width = 300 if c == "Key" else 160
            anchor = "w" if c in ("Key", "Presence") else "e"
            self.tree.column(c, width=width, anchor=anchor)

        yscroll = ttk.Scrollbar(results, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(results, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        results.grid_rowconfigure(0, weight=1)
        results.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("even", background=self.theme.colors["tree_bg"])
        self.tree.tag_configure("odd", background=self.theme.colors["tree_alt"])

        self.status = ttk.Label(self, text="Ready", foreground=self.theme.colors["accent"])
        self.status.grid(row=2, column=0, sticky="w", pady=(6, 0))

    def set_status(self, msg: str):
        self.status.configure(text=msg)

    def clear(self):
        self.tree.delete(*self.tree.get_children())

    def populate(self, rows: List[tuple]):
        self.clear()
        count = 0
        for key, whq, ecq, diff, pres in rows:
            tag = "odd" if count % 2 else "even"
            self.tree.insert("", "end",
                             values=(key, f"{whq:.0f}", f"{ecq:.0f}", f"{diff:.0f}", pres),
                             tags=(tag,))
            count += 1
        self.set_status(f"Rows: {count:,}")

    def filtered(self) -> List[tuple]:
        data = []
        for iid in self.tree.get_children():
            key, wq, eq, diff, pres = self.tree.item(iid, "values")
            data.append((key, float(str(wq).replace(",", "")), float(str(eq).replace(",", "")), float(str(diff).replace(",", "")), pres))
        only_d = self.only_diff.get()
        pres = self.presence.get()
        op = self.diff_op.get()
        thr = float(self.diff_thr.get() or 0.0)

        def ok(row):
            _, _, _, d, p = row
            if pres != "All" and p != pres:
                return False
            if only_d and abs(d) == 0.0:
                return False
            if op == "!=" and not (d != thr): return False
            if op == ">"  and not (d >  thr): return False
            if op == "<"  and not (d <  thr): return False
            if op == ">=" and not (d >= thr): return False
            if op == "<=" and not (d <= thr): return False
            if op == "==" and not (d == thr): return False
            return True

        return [r for r in data if ok(r)]

    def _sort_tree_by(self, col_name: str, descending: bool):
        rows = [(self.tree.set(k, col_name), k) for k in self.tree.get_children("")]
        def _to_num(v):
            try:
                return float(str(v).replace(",", ""))
            except Exception:
                return None
        numeric = {"Warehouse Qty", "Ecommerce Qty", "Difference"}
        if col_name in numeric:
            decorated = [((_to_num(v) if _to_num(v) is not None else float("inf")), k) for v, k in rows]
        else:
            decorated = [(str(v).lower(), k) for v, k in rows]
        decorated.sort(reverse=descending)
        for i, (_, k) in enumerate(decorated):
            self.tree.move(k, "", i)
        self.tree.heading(col_name, command=lambda: self._sort_tree_by(col_name, not descending))


class CompositeTab(ttk.Frame):
    def __init__(self, parent, theme: ThemeManager):
        super().__init__(parent, padding=10)
        self.theme = theme
        self.engine = create_engine("sqlite+pysqlite:///:memory:", future=True)

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=2)
        self.columnconfigure(2, weight=3)

        left = ttk.Frame(self, padding=10)
        left.grid(row=0, column=0, sticky="nsew")
        self.wh_card = SourceCard(left, "Warehouse", theme)
        self.ec_card = SourceCard(left, "Ecommerce", theme)
        self.wh_card.grid(row=0, column=0, sticky="ew")
        self.ec_card.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        left.grid_rowconfigure(2, weight=1)

        prof = ttk.Labelframe(left, text="Profiles", padding=10)
        prof.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        prof.grid_columnconfigure(1, weight=1)
        ttk.Label(prof, text="Name:").grid(row=0, column=0, sticky="w")
        self.profile_name = tk.StringVar()
        ttk.Entry(prof, textvariable=self.profile_name).grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.profiles = load_profiles()
        ttk.Label(prof, text="Load:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.profile_pick = tk.StringVar()
        self.profile_dd = ttk.Combobox(prof, textvariable=self.profile_pick, state="readonly",
                                       values=sorted(self.profiles.keys()))
        self.profile_dd.grid(row=1, column=1, sticky="ew", padx=(6, 6), pady=(6, 0))
        ttk.Button(prof, text="Save Current", command=self._save_profile).grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(prof, text="Load Selected", command=self._load_profile).grid(row=2, column=1, sticky="ew", padx=(6, 0), pady=(8, 0))

        self.results = CompositeResultsPane(self, theme)
        self.results.grid(row=0, column=2, sticky="nsew", padx=(6, 10), pady=10)

        self.results.btn_compare.configure(command=self._compare)
        self.results.btn_export.configure(command=self._export)

        self.bind_all("<Control-Return>", lambda e: self._compare())
        self.bind_all("<Control-s>", lambda e: self._export())

    def _save_profile(self):
        name = self.profile_name.get().strip()
        if not name:
            messagebox.showwarning("Profiles", "Enter a profile name.")
            return
        cfg = {
            "wh": {
                "keys": self.wh_card.get_key_cols(),
                "qty": self.wh_card.get_qty_col(),
                "norm": self.wh_card.get_norm(),
            },
            "ec": {
                "keys": self.ec_card.get_key_cols(),
                "qty": self.ec_card.get_qty_col(),
                "norm": self.ec_card.get_norm(),
            },
            "filters": {
                "only_diff": bool(self.results.only_diff.get()),
                "presence": self.results.presence.get(),
                "diff_op": self.results.diff_op.get(),
                "diff_thr": float(self.results.diff_thr.get() or 0.0),
            },
        }
        self.profiles[name] = cfg
        save_profiles(self.profiles)
        self.profile_dd["values"] = sorted(self.profiles.keys())
        self.profile_pick.set(name)
        messagebox.showinfo("Profiles", f"Saved '{name}'.")

    def _load_profile(self):
        name = self.profile_pick.get().strip()
        if not name or name not in self.profiles:
            messagebox.showwarning("Profiles", "Pick a saved profile to load.")
            return
        cfg = self.profiles[name]
        wh, ec = cfg.get("wh", {}), cfg.get("ec", {})
        if self.wh_card.df is not None:
            self.wh_card.chips.set_values([c for c in wh.get("keys", []) if c in self.wh_card.df.columns])
            if wh.get("qty") in (self.wh_card.df.columns if self.wh_card.df is not None else []):
                self.wh_card.qty_var.set(wh["qty"])
        if self.ec_card.df is not None:
            self.ec_card.chips.set_values([c for c in ec.get("keys", []) if c in self.ec_card.df.columns])
            if ec.get("qty") in (self.ec_card.df.columns if self.ec_card.df is not None else []):
                self.ec_card.qty_var.set(ec["qty"])
        for card, norm in ((self.wh_card, wh.get("norm", {})), (self.ec_card, ec.get("norm", {}))):
            if norm:
                card.case_var.set(norm.get("case", "lower"))
                card.trim_var.set(bool(norm.get("trim", True)))
                card.collapse_var.set(bool(norm.get("collapse_spaces", True)))
                card.pad_var.set(int(norm.get("zero_pad", 0)))
                card._refresh_preview()
        f = cfg.get("filters", {})
        self.results.only_diff.set(bool(f.get("only_diff", False)))
        self.results.presence.set(f.get("presence", "All"))
        self.results.diff_op.set(f.get("diff_op", "!="))
        self.results.diff_thr.set(float(f.get("diff_thr", 0.0)))
        messagebox.showinfo("Profiles", f"Loaded '{name}'.")

    def _compare(self):
        self.results.set_status("Comparing...")
        self.results.clear()

        if self.wh_card.df is None or self.ec_card.df is None:
            messagebox.showwarning("Compare", "Load both Warehouse and Ecommerce CSVs first.")
            self.results.set_status("Ready")
            return

        wh_keys = self.wh_card.get_key_cols()
        ec_keys = self.ec_card.get_key_cols()
        wh_qty = self.wh_card.get_qty_col()
        ec_qty = self.ec_card.get_qty_col()
        if not wh_keys or not ec_keys or not wh_qty or not ec_qty:
            messagebox.showwarning("Compare", "Set composite keys and quantities on both sides.")
            self.results.set_status("Ready")
            return

        wh_norm_cfg = self.wh_card.get_norm()
        ec_norm_cfg = self.ec_card.get_norm()

        wh = self.wh_card.df.copy()
        ec = self.ec_card.df.copy()
        wh["_key"] = build_composite_key_series(wh, wh_keys, wh_norm_cfg)
        ec["_key"] = build_composite_key_series(ec, ec_keys, ec_norm_cfg)
        wh = wh[wh["_key"] != ""].drop_duplicates(subset=["_key"], keep="last")
        ec = ec[ec["_key"] != ""].drop_duplicates(subset=["_key"], keep="last")

        if wh.empty and ec.empty:
            messagebox.showinfo("Compare", "No valid keys after normalization.")
            self.results.set_status("Ready")
            return

        wh_norm = pd.DataFrame({"key": wh["_key"].values, "qty": coerce_numeric(wh[wh_qty]).values})
        ec_norm = pd.DataFrame({"key": ec["_key"].values, "qty": coerce_numeric(ec[ec_qty]).values})

        with self.engine.begin() as conn:
            conn.exec_driver_sql("DROP TABLE IF EXISTS wh_norm;")
            conn.exec_driver_sql("DROP TABLE IF EXISTS ec_norm;")
            wh_norm.to_sql("wh_norm", conn, index=False)
            ec_norm.to_sql("ec_norm", conn, index=False)
            conn.exec_driver_sql("CREATE INDEX IF NOT EXISTS idx_wh_key ON wh_norm(key);")
            conn.exec_driver_sql("CREATE INDEX IF NOT EXISTS idx_ec_key ON ec_norm(key);")

            left_sql = """
                SELECT
                    COALESCE(w.key, e.key) AS key,
                    COALESCE(w.qty, 0.0)   AS wh_qty,
                    COALESCE(e.qty, 0.0)   AS ec_qty,
                    CASE
                        WHEN e.key IS NULL THEN 'Only in Warehouse'
                        WHEN w.key IS NULL THEN 'Only in Ecommerce'
                        ELSE 'Both'
                    END AS presence
                FROM wh_norm w
                LEFT JOIN ec_norm e ON w.key = e.key
            """
            right_only_sql = """
                SELECT
                    COALESCE(w.key, e.key) AS key,
                    COALESCE(w.qty, 0.0)   AS wh_qty,
                    COALESCE(e.qty, 0.0)   AS ec_qty,
                    CASE
                        WHEN e.key IS NULL THEN 'Only in Warehouse'
                        WHEN w.key IS NULL THEN 'Only in Ecommerce'
                        ELSE 'Both'
                    END AS presence
                FROM ec_norm e
                LEFT JOIN wh_norm w ON w.key = e.key
                WHERE w.key IS NULL
            """
            union_sql = f"""
                WITH all_rows AS (
                    {left_sql}
                    UNION ALL
                    {right_only_sql}
                )
                SELECT key,
                       wh_qty,
                       ec_qty,
                       (wh_qty - ec_qty) AS diff,
                       presence
                FROM all_rows
            """
            rows = conn.execute(text(union_sql)).fetchall()

        rows = self._apply_filters(rows)
        self.results.populate(rows)

    def _apply_filters(self, rows: List[tuple]) -> List[tuple]:
        only_d = self.results.only_diff.get()
        pres = self.results.presence.get()
        op = self.results.diff_op.get()
        thr = float(self.results.diff_thr.get() or 0.0)

        def ok(row):
            _, _, _, d, p = row
            if pres != "All" and p != pres:
                return False
            if only_d and abs(d) == 0.0:
                return False
            if op == "!=" and not (d != thr): return False
            if op == ">"  and not (d >  thr): return False
            if op == "<"  and not (d <  thr): return False
            if op == ">=" and not (d >= thr): return False
            if op == "<=" and not (d <= thr): return False
            if op == "==" and not (d == thr): return False
            return True

        return [r for r in rows if ok(r)]

    def _export(self):
        if not self.results.tree.get_children():
            messagebox.showwarning("Export", "Nothing to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Save results as CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return
        cols = self.results.tree["columns"]
        data = []
        for iid in self.results.tree.get_children():
            data.append(self.results.tree.item(iid, "values"))
        df = pd.DataFrame(data, columns=cols)
        try:
            df.to_csv(path, index=False)
            messagebox.showinfo("Export", f"Saved: {path}")
        except Exception as e:
            messagebox.showerror("Export", f"Failed to save:\n{e}")

class MainApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Inventory Qty Tools — Checker + Composite-Key Compare")
        self.root.minsize(1100, 600)

        self.theme = ThemeManager(self.root)
        self.cfg = load_config()
        self.ctx = ContextMenuManager(self.root)

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)

        self.checker_tab = CheckerTab(nb, self.theme, self.cfg)
        nb.add(self.checker_tab, text="Single-SKU Checker")

        self.composite_tab = CompositeTab(nb, self.theme)
        nb.add(self.composite_tab, text="Composite-Key Compare")

        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    if sys.platform.startswith("linux") and not sys.stdout.isatty():
        print("Run in a desktop session.")
    else:
        MainApp().run()

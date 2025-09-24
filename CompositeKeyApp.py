#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Composite-Key Stock Compare — Tkinter + SQLAlchemy (SQLite in-memory)
Compact UI edition
"""

import csv
import json
import re
import sys
from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd
from sqlalchemy import create_engine, text

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ----------------------------- THEME / COLORS -----------------------------
APP_BG = "#0e1018"
PANEL_BG = "#121424"
FG = "#e6e6e6"
ACCENT = "#00d4ff"
TREE_BG = "#151826"
TREE_SEL = "#1e2a52"
ROW_STRIPE = "#0f1430"
BORDER = "#3a3f55"

PROFILES_FILE = "ck_profiles.json"


# ----------------------------- UTILITIES -----------------------------
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
        parts.append(
            df[c].astype(str).map(
                lambda x: normalize_piece(
                    x,
                    case=norm["case"],
                    trim=norm["trim"],
                    collapse_spaces=norm["collapse_spaces"],
                    zero_pad=norm["zero_pad"],
                )
            )
        )
    out = parts[0]
    for p in parts[1:]:
        out = out + "|" + p
    return out


def coerce_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0).astype(float)


# ----------------------------- COMPACT WIDGETS -----------------------------
class Collapsible(ttk.Frame):
    def __init__(self, parent, title="Advanced"):
        super().__init__(parent)
        self._open = False
        self.btn = ttk.Button(self, text=f"▸ {title}", style="TButton", command=self._toggle, width=18)
        self.btn.grid(row=0, column=0, sticky="w")
        self.body = ttk.Frame(self)
        self.grid_columnconfigure(0, weight=1)

    def _toggle(self):
        self._open = not self._open
        self.btn.configure(text=("▾ " if self._open else "▸ ") + self.btn.cget("text")[2:])
        if self._open:
            self.body.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        else:
            self.body.grid_forget()


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
            chip = ttk.Frame(self, padding=(6, 1), style="Chip.TFrame")
            ttk.Label(chip, text=col, style="Chip.TLabel").pack(side="left")
            ttk.Button(chip, text="×", width=2, style="Chip.TButton",
                       command=lambda c=col: self.remove(c)).pack(side="left", padx=(6, 0))
            chip.pack(side="left", padx=3)


# ----------------------------- SOURCE CARD (COMPACT) -----------------------------
class SourceCard(ttk.Labelframe):
    def __init__(self, parent, title: str):
        super().__init__(parent, text=title, padding=(8, 8), style="TLabelframe")
        self.df: Optional[pd.DataFrame] = None
        self.title = title

        row = 0
        self.btn = ttk.Button(self, text="Load CSV", command=self._on_load, width=14)
        self.btn.grid(row=row, column=0, sticky="w")
        self.badge = ttk.Label(self, text="No file", foreground=ACCENT)
        self.badge.grid(row=row, column=1, sticky="e")
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        row += 1

        ttk.Label(self, text="Add key column").grid(row=row, column=0, sticky="w", pady=(8, 2))
        self.col_var = tk.StringVar()
        self.col_dd = ttk.Combobox(self, textvariable=self.col_var, state="readonly", width=28)
        self.col_dd.grid(row=row, column=1, sticky="e", pady=(8, 2))
        row += 1

        small = ttk.Frame(self)
        small.grid(row=row, column=0, columnspan=2, sticky="ew")
        self.add_btn = ttk.Button(small, text="Add →", command=self._add_selected, width=10)
        self.add_btn.pack(side="left")
        ttk.Label(small, text="Qty").pack(side="left", padx=(10, 4))
        self.qty_var = tk.StringVar()
        self.qty_dd = ttk.Combobox(small, textvariable=self.qty_var, state="readonly", width=20)
        self.qty_dd.pack(side="left")
        row += 1

        ttk.Label(self, text="Composite Key").grid(row=row, column=0, sticky="w", pady=(8, 2))
        row += 1

        self.chips = ChipBar(self, on_change=lambda cols: self._refresh_preview())
        self.chips.grid(row=row, column=0, columnspan=2, sticky="w")
        row += 1

        adv = Collapsible(self, "Normalization & Preview")
        adv.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        row += 1

        norm = adv.body
        ttk.Label(norm, text="Case").grid(row=0, column=0, sticky="w")
        self.case_var = tk.StringVar(value="lower")
        ttk.Combobox(norm, textvariable=self.case_var, state="readonly", width=8,
                     values=["lower", "upper", "as-is"]).grid(row=0, column=1, sticky="w", padx=(6, 18))
        self.trim_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(norm, text="Trim", variable=self.trim_var).grid(row=0, column=2, sticky="w")
        self.collapse_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(norm, text="Collapse spaces", variable=self.collapse_var).grid(row=0, column=3, sticky="w", padx=(8, 0))
        ttk.Label(norm, text="Zero-pad").grid(row=0, column=4, sticky="w", padx=(12, 0))
        self.pad_var = tk.IntVar(value=0)
        ttk.Spinbox(norm, from_=0, to=12, textvariable=self.pad_var, width=4).grid(row=0, column=5, sticky="w", padx=(6, 0))

        ttk.Label(norm, text="Preview (10)").grid(row=1, column=0, sticky="w", pady=(8, 2))
        self.preview = tk.Text(norm, height=4, width=40, bg=TREE_BG, fg=FG, relief="flat")
        self.preview.grid(row=2, column=0, columnspan=6, sticky="ew")
        for i in range(6):
            norm.grid_columnconfigure(i, weight=1)

    def _add_selected(self):
        col = self.col_var.get()
        if not col:
            return
        self.chips.add(col)
        self._refresh_preview()

    def _on_load(self):
        fpath = filedialog.askopenfilename(
            title=f"Select {self.title} CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not fpath:
            return
        p = Path(fpath)
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
        self.badge.configure(text=f"{len(df):,} rows · {len(cols)} cols")
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


# ----------------------------- RESULTS PANE (COMPACT) -----------------------------
class ResultsPane(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=(8, 8))
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        bar.grid_columnconfigure(0, weight=1)
        bar.grid_columnconfigure(1, weight=1)
        bar.grid_columnconfigure(2, weight=0)

        left = ttk.Frame(bar)
        left.grid(row=0, column=0, sticky="w")
        self.only_diff = tk.BooleanVar(value=False)
        ttk.Checkbutton(left, text="Only differences", variable=self.only_diff).pack(side="left")
        ttk.Label(left, text="Presence").pack(side="left", padx=(10, 4))
        self.presence = tk.StringVar(value="All")
        ttk.Combobox(left, textvariable=self.presence, state="readonly",
                     values=["All", "Both", "Only in Warehouse", "Only in Ecommerce"], width=18).pack(side="left")

        mid = ttk.Frame(bar)
        mid.grid(row=0, column=1, sticky="w")
        ttk.Label(mid, text="Diff").pack(side="left")
        self.diff_op = tk.StringVar(value="!=")
        ttk.Combobox(mid, textvariable=self.diff_op, state="readonly",
                     values=["!=", ">", "<", ">=", "<=", "=="], width=4).pack(side="left", padx=(6, 4))
        self.diff_thr = tk.DoubleVar(value=0.0)
        ttk.Entry(mid, textvariable=self.diff_thr, width=8).pack(side="left")

        right = ttk.Frame(bar)
        right.grid(row=0, column=2, sticky="e")
        self.btn_compare = ttk.Button(right, text="Compare", width=12)
        self.btn_export = ttk.Button(right, text="Export", width=10)
        self.btn_compare.pack(side="left", padx=(0, 6))
        self.btn_export.pack(side="left")

        table = ttk.Frame(self)
        table.grid(row=1, column=0, sticky="nsew")
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        cols = ("Key", "Warehouse Qty", "Ecommerce Qty", "Difference", "Presence")
        self.tree = ttk.Treeview(table, columns=cols, show="headings", style="Dark.Treeview")
        for c in cols:
            self.tree.heading(c, text=c, command=lambda col=c: self._sort_tree_by(col, False))
            width = 360 if c == "Key" else 120
            anchor = "w" if c in ("Key", "Presence") else "e"
            self.tree.column(c, width=width, anchor=anchor)

        yscroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table.grid_rowconfigure(0, weight=1)
        table.grid_columnconfigure(0, weight=1)

        self.status = ttk.Label(self, text="Ready", foreground=ACCENT)
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
        self.tree.tag_configure("even", background=TREE_BG)
        self.tree.tag_configure("odd", background=ROW_STRIPE)
        self.set_status(f"Rows: {count:,}")

    def filtered(self) -> List[tuple]:
        data = []
        for iid in self.tree.get_children():
            key, wq, eq, diff, pres = self.tree.item(iid, "values")
            data.append((key, float(wq.replace(",", "")), float(eq.replace(",", "")), float(diff.replace(",", "")), pres))
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


# ----------------------------- PROFILES -----------------------------
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


# ----------------------------- MAIN APP -----------------------------
class CompositeKeyApp(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master)
        self.master = master
        self._init_style()
        self.pack(fill="both", expand=True)

        self.engine = create_engine("sqlite+pysqlite:///:memory:", future=True)

        root_grid = ttk.Frame(self, padding=6)
        root_grid.pack(fill="both", expand=True)
        root_grid.grid_columnconfigure(0, weight=0)
        root_grid.grid_columnconfigure(1, weight=1)
        root_grid.grid_rowconfigure(0, weight=1)

        # LEFT SIDEBAR (compact)
        left = ttk.Frame(root_grid)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 6))
        self.wh_card = SourceCard(left, "Warehouse")
        self.ec_card = SourceCard(left, "Ecommerce")
        self.wh_card.pack(fill="x")
        self.ec_card.pack(fill="x", pady=(6, 0))

        prof = ttk.Labelframe(left, text="Profiles", padding=(8, 8))
        prof.pack(fill="x", pady=(6, 0))
        prof.grid_columnconfigure(1, weight=1)
        ttk.Label(prof, text="Name").grid(row=0, column=0, sticky="w")
        self.profile_name = tk.StringVar()
        ttk.Entry(prof, textvariable=self.profile_name, width=18).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.profiles = load_profiles()
        ttk.Label(prof, text="Load").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.profile_pick = tk.StringVar()
        self.profile_dd = ttk.Combobox(prof, textvariable=self.profile_pick, state="readonly",
                                       values=sorted(self.profiles.keys()), width=18)
        self.profile_dd.grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))
        small_btns = ttk.Frame(prof)
        small_btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(small_btns, text="Save", command=self._save_profile, width=10).pack(side="left")
        ttk.Button(small_btns, text="Load", command=self._load_profile, width=10).pack(side="left", padx=(6, 0))

        # RIGHT MAIN
        main = ttk.Frame(root_grid)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)

        self.results = ResultsPane(main)
        self.results.grid(row=0, column=0, sticky="nsew")

        self.results.btn_compare.configure(command=self._compare)
        self.results.btn_export.configure(command=self._export)

        master.bind("<Control-Return>", lambda e: self._compare())
        master.bind("<Control-s>", lambda e: self._export())

    # THEME/STYLES
    def _init_style(self):
        self.master.title("Composite-Key Stock Compare — Compact UI")
        self.master.geometry("1180x720")
        self.master.configure(bg=APP_BG)
        style = ttk.Style(self.master)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Ttk core
        style.configure("TFrame", background=APP_BG)
        style.configure("TLabelframe", background=PANEL_BG, foreground=FG, bordercolor=BORDER)
        style.configure("TLabelframe.Label", background=PANEL_BG, foreground=FG, font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background=APP_BG, foreground=FG, font=("Segoe UI", 10))
        style.configure("TButton", background="#2a2f42", foreground=FG, font=("Segoe UI", 10), padding=(8, 4))
        style.map("TButton", background=[("active", "#333a55")])
        style.configure("TCheckbutton", background=APP_BG, foreground=FG)
        style.configure("TEntry", fieldbackground="#1b2137", foreground=FG, background="#1b2137")
        style.configure("TSpinbox", fieldbackground="#1b2137", foreground=FG, arrowsize=14)

        # Combobox (including dropdown list colors via option_add)
        style.configure("TCombobox",
                        fieldbackground="#1b2137",
                        background="#1b2137",
                        foreground=FG,
                        arrowcolor=FG)
        style.map("TCombobox",
                  fieldbackground=[("readonly", "#1b2137")],
                  foreground=[("readonly", FG)])
        # Force dark dropdown listbox
        self.master.option_add("*TCombobox*Listbox*Background", "#0f1430")
        self.master.option_add("*TCombobox*Listbox*Foreground", FG)
        self.master.option_add("*TCombobox*Listbox*selectBackground", TREE_SEL)
        self.master.option_add("*TCombobox*Listbox*selectForeground", FG)
        self.master.option_add("*Menu*Background", "#0f1430")
        self.master.option_add("*Menu*Foreground", FG)
        self.master.option_add("*Menu*activeBackground", TREE_SEL)
        self.master.option_add("*Menu*activeForeground", FG)
        self.master.option_add("*Menu*borderColor", BORDER)
        self.master.option_add("*Menu*relief", "flat")
        self.master.option_add("*Menu*font", ("Segoe UI", 10))
        self.master.option_add("*TCombobox*Listbox*font", ("Segoe UI", 10))
        self.master.option_add("*TCombobox*Listbox*borderWidth", 0)
        self.master.option_add("*TCombobox*Listbox*highlightThickness", 0)
        self.master.option_add("*TCombobox*Listbox*selectBorderWidth", 0)
        self.master.option_add("*TCombobox*Listbox*highlightColor", BORDER)
        self.master.option_add("*TCombobox*Listbox*highlightBackground", BORDER)
        self.master.option_add("*TCombobox*Listbox*borderColor", BORDER)
        self.master.option_add("*TCombobox*Listbox*relief", "flat")
        self.master.option_add("*TCombobox*Listbox*selectRelief", "flat")
        self.master.option_add("*TCombobox*Listbox*selectBorderColor", BORDER)
        self.master.option_add("*TCombobox*Listbox*selectBorderWidth", 0)
        self.master.option_add("*TCombobox*Listbox*selectHighlightThickness", 0)
        self.master.option_add("*TCombobox*Listbox*selectHighlightColor", BORDER)
        self.master.option_add("*TCombobox*Listbox*selectHighlightBackground", BORDER)
        self.master.option_add("*TCombobox*Listbox*font", ("Segoe UI", 10))
        self.master.option_add("*TCombobox*Listbox*padding", 0)
        self.master.option_add("*TCombobox*Listbox*cursor", "arrow")
        self.master.option_add("*TCombobox*Listbox*selectMode", "browse")
        self.master.option_add("*TCombobox*Listbox*width", 0)
        self.master.option_add("*TCombobox*Listbox*height", 0)
        self.master.option_add("*TCombobox*Listbox*takeFocus", 0)
        self.master.option_add("*TCombobox*Listbox*exportselection", 0)
        self.master.option_add("*TCombobox*Listbox*activestyle", "none")
        self.master.option_add("*TCombobox*Listbox*selectBackground", TREE_SEL)
        self.master.option_add("*TCombobox*Listbox*selectForeground", FG)
        self.master.option_add("*TCombobox*Listbox*selectPadding", 0)
        self.master.option_add("*TCombobox*Listbox*selectCursor", "arrow")
        self.master.option_add("*TCombobox*Listbox*selectMode", "browse")
        self.master.option_add("*TCombobox*Listbox*selectWidth", 0)
        self.master.option_add("*TCombobox*Listbox*selectHeight", 0)
        self.master.option_add("*TCombobox*Listbox*selectTakeFocus", 0)
        self.master.option_add("*TCombobox*Listbox*selectExportselection", 0)
        self.master.option_add("*TCombobox*Listbox*selectActivestyle", "none")
        self.master.option_add("*TCombobox*Listbox*selectFont", ("Segoe UI", 10, "bold"))
 



        style.configure("Dark.Treeview",
                        background=TREE_BG,
                        foreground=FG,
                        fieldbackground=TREE_BG,
                        bordercolor=BORDER,
                        rowheight=24)
        style.map("Dark.Treeview",
                  background=[("selected", TREE_SEL)],
                  foreground=[("selected", FG)])
        style.configure("Dark.Treeview.Heading",
                        background="#20263f",
                        foreground=FG,
                        relief="flat",
                        font=("Segoe UI", 10, "bold"))
        style.map("Dark.Treeview.Heading", background=[("active", "#263054")])

        # Chips
        style.configure("Chip.TFrame", background="#1f2540", borderwidth=0)
        style.configure("Chip.TLabel", background="#1f2540", foreground=FG, font=("Segoe UI", 9))
        style.configure("Chip.TButton", background="#2a2f42", foreground=FG, padding=0, width=2)

    # PROFILES
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
        profs = load_profiles()
        profs[name] = cfg
        save_profiles(profs)
        self.results.set_status(f"Saved profile '{name}'")
        self.profile_dd["values"] = sorted(profs.keys())
        self.profile_pick.set(name)

    def _load_profile(self):
        profs = load_profiles()
        name = self.profile_pick.get().strip()
        if not name or name not in profs:
            messagebox.showwarning("Profiles", "Pick a saved profile to load.")
            return
        cfg = profs[name]
        wh, ec = cfg.get("wh", {}), cfg.get("ec", {})
        if self.wh_card.df is not None:
            self.wh_card.chips.set_values([c for c in wh.get("keys", []) if c in self.wh_card.df.columns])
            if wh.get("qty") in self.wh_card.df.columns:
                self.wh_card.qty_var.set(wh["qty"])
        if self.ec_card.df is not None:
            self.ec_card.chips.set_values([c for c in ec.get("keys", []) if c in self.ec_card.df.columns])
            if ec.get("qty") in self.ec_card.df.columns:
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
        self.results.set_status(f"Loaded profile '{name}'")

    # COMPARE / EXPORT
    def _compare(self):
        self.results.set_status("Comparing…")
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
            messagebox.showwarning("Compare", "Set composite keys and quantity columns on both sides.")
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
        data = [self.results.tree.item(iid, "values") for iid in self.results.tree.get_children()]
        df = pd.DataFrame(data, columns=cols)
        try:
            df.to_csv(path, index=False)
            messagebox.showinfo("Export", f"Saved: {path}")
        except Exception as e:
            messagebox.showerror("Export", f"Failed to save:\n{e}")


def main():
    root = tk.Tk()
    app = CompositeKeyApp(root)
    root.mainloop()


if __name__ == "__main__":
    if sys.platform.startswith("linux") and not sys.stdout.isatty():
        print("Run in a desktop session.")
    else:
        main()

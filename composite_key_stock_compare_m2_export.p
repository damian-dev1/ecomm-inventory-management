#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import csv
import re
import math
import json
import traceback
from dataclasses import dataclass, asdict, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import pandas as pd
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = False

from ttkbootstrap import Window, Style, ttk
from ttkbootstrap.constants import *

APP_NAME = "Composite Key Stock Compare — Magento 2 Export"
DEFAULT_THEME = "cyborg"
PROFILE_EXT = ".m2profile.json"

@dataclass
class Normalization:
    trim: bool = True
    casefold: bool = True
    collapse_ws: bool = True
    zero_pad_width: int = 0

@dataclass
class Mapping:
    file_path: Optional[str] = None
    columns: List[str] = field(default_factory=list)
    key_order: List[str] = field(default_factory=list)           # ordered columns for composite key
    qty_col: Optional[str] = None                                # quantity column
    sku_col: Optional[str] = None                                # which column writes to Magento 'sku'
    normalization: Normalization = field(default_factory=Normalization)

@dataclass
class ExportCfg:
    out_dir: str = str(Path.home() / "Desktop" / "m2_exports")
    chunk_size: int = 10_000
    source_codes: List[str] = None

    def __post_init__(self):
        if self.source_codes is None:
            self.source_codes = ["pos_337", "src_virtualstock"]

@dataclass
class AutomationCfg:
    enabled: bool = False
    autorun_on_start: bool = False
    poll_seconds: int = 30
    watch_ec_path: str = ""
    watch_wh_path: str = ""
    run_if_changed: bool = True
    _last_mtime_ec: float = 0.0
    _last_mtime_wh: float = 0.0

def read_csv_columns(path: str) -> List[str]:
    if _HAS_PANDAS:
        df = pd.read_csv(path, nrows=1)
        return list(df.columns)
    else:
        with open(path, "r", newline="", encoding="utf-8", errors="replace") as f:
            reader = csv.reader(f)
            header = next(reader, [])
        return header

def count_csv_rows(path: str) -> int:
    # Efficient line count (minus header)
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return max(0, sum(1 for _ in f) - 1)
    except Exception:
        return 0

def load_rows(path: str, usecols: List[str]) -> List[Dict]:
    if _HAS_PANDAS:
        df = pd.read_csv(path, usecols=usecols, dtype=str, keep_default_na=False)
        return df.to_dict(orient="records")
    else:
        with open(path, "r", newline="", encoding="utf-8", errors="replace") as f:
            reader = csv.DictReader(f)
            rows = []
            for row in reader:
                rows.append({c: str(row.get(c, "")) for c in usecols})
            return rows

def normalize_value(v: str, norm: Normalization) -> str:
    s = "" if v is None else str(v)
    if norm.trim:
        s = s.strip()
    if norm.collapse_ws:
        s = re.sub(r"\s+", " ", s)
    if norm.casefold:
        s = s.casefold()
    if norm.zero_pad_width and s.isdigit():
        s = s.zfill(norm.zero_pad_width)
    return s

def build_composite(row: Dict[str, str], key_order: List[str], norm: Normalization, sep: str = "⋮") -> str:
    parts = [normalize_value(row.get(col, ""), norm) for col in key_order]
    return sep.join(parts)

def to_int_safe(v: str) -> int:
    try:
        return int(float(str(v).strip()))
    except Exception:
        return 0

class CompositeKeyDashboard:
    def __init__(self, root: Window):
        self.root = root
        self.root.title(f"🛠️ {APP_NAME}")
        self.root.geometry("1320x840")
        self.root.minsize(1160, 760)

        # state
        self.map_ec = Mapping(columns=[], key_order=[])
        self.map_wh = Mapping(columns=[], key_order=[])
        self.export_cfg = ExportCfg()
        self.auto_cfg = AutomationCfg()
        self.current_preview_side = tk.StringVar(value="ECommerce")

        # vars
        self.var_out_dir = tk.StringVar(value=self.export_cfg.out_dir)
        self.var_chunk = tk.IntVar(value=self.export_cfg.chunk_size)

        # automation vars
        self.var_auto_enabled = tk.BooleanVar(value=self.auto_cfg.enabled)
        self.var_auto_autorun = tk.BooleanVar(value=self.auto_cfg.autorun_on_start)
        self.var_auto_seconds = tk.IntVar(value=self.auto_cfg.poll_seconds)
        self.var_auto_run_if_changed = tk.BooleanVar(value=self.auto_cfg.run_if_changed)
        self.var_watch_ec = tk.StringVar(value=self.auto_cfg.watch_ec_path)
        self.var_watch_wh = tk.StringVar(value=self.auto_cfg.watch_wh_path)

        # profile path
        self.var_profile_path = tk.StringVar(value="")

        # automation scheduling
        self._auto_job_id: Optional[str] = None
        self._auto_running = False
        self._next_tick_eta: Optional[datetime] = None

        # stats/state vars (sidebar)
        self.stat_ec_rows = tk.IntVar(value=0)
        self.stat_wh_rows = tk.IntVar(value=0)
        self.stat_ec_keys = tk.StringVar(value="not set")
        self.stat_wh_keys = tk.StringVar(value="not set")
        self.stat_last_run = tk.StringVar(value="–")
        self.stat_next_run = tk.StringVar(value="–")
        self.stat_last_diff = tk.IntVar(value=0)
        self.stat_output_dir = tk.StringVar(value=self.var_out_dir.get())
        self.stat_auto_state = tk.StringVar(value="Stopped")

        # build ui
        self._build_ui()
        self._bind_shortcuts()
        self._log("Ready")

        if self.var_auto_enabled.get() and self.var_auto_autorun.get():
            self._start_automation()

    # ---------------- UI Layout ----------------
    def _build_ui(self):
        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Sidebar (Files, Stats/State, Export)
        sidebar = ttk.Frame(self.root, padding=12, width=380)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=0, column=1, sticky="nsew")
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        # Files
        self._build_file_block(sidebar)
        ttk.Separator(sidebar).pack(fill=X, pady=10)

        # Stats & State (replaces composite key config in sidebar)
        self._build_stats_block(sidebar)
        ttk.Separator(sidebar).pack(fill=X, pady=10)

        # Export section (kept in sidebar)
        self._build_export_block(sidebar)

        # Main tabs / preview / logs
        self._build_main_area(main)

        # Status bar
        self.status = ttk.Label(self.root, text="Ready", anchor="w", bootstyle=SECONDARY)
        self.status.grid(row=1, column=0, columnspan=2, sticky="ew")

    def _build_file_block(self, parent: ttk.Frame):
        f = ttk.LabelFrame(parent, text="Files", padding=10)
        f.pack(fill=X)

        ttk.Label(f, text="ECommerce CSV").pack(anchor="w")
        row1 = ttk.Frame(f); row1.pack(fill=X, pady=(2, 6))
        self.ec_entry = ttk.Entry(row1); self.ec_entry.pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row1, text="Browse", command=self._pick_ec).pack(side=LEFT, padx=6)

        ttk.Label(f, text="Warehouse CSV").pack(anchor="w")
        row2 = ttk.Frame(f); row2.pack(fill=X, pady=(2, 2))
        self.wh_entry = ttk.Entry(row2); self.wh_entry.pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row2, text="Browse", command=self._pick_wh).pack(side=LEFT, padx=6)

        ttk.Button(f, text="Load Columns", bootstyle=INFO, command=self._load_columns).pack(fill=X, pady=(8, 0))

    def _build_stats_block(self, parent: ttk.Frame):
        lf = ttk.LabelFrame(parent, text="Stats & State", padding=10)
        lf.pack(fill=X)

        # Row counts
        g1 = ttk.Frame(lf); g1.pack(fill=X)
        ttk.Label(g1, text="ECommerce rows").grid(row=0, column=0, sticky="w")
        ttk.Label(g1, textvariable=self.stat_ec_rows, bootstyle=INFO).grid(row=0, column=1, sticky="e")
        ttk.Label(g1, text="Warehouse rows").grid(row=1, column=0, sticky="w", pady=(4,0))
        ttk.Label(g1, textvariable=self.stat_wh_rows, bootstyle=INFO).grid(row=1, column=1, sticky="e", pady=(4,0))

        ttk.Separator(lf).pack(fill=X, pady=8)

        # Key config state
        g2 = ttk.Frame(lf); g2.pack(fill=X)
        ttk.Label(g2, text="EC key columns").grid(row=0, column=0, sticky="w")
        ttk.Label(g2, textvariable=self.stat_ec_keys, bootstyle=SECONDARY).grid(row=0, column=1, sticky="e")
        ttk.Label(g2, text="WH key columns").grid(row=1, column=0, sticky="w", pady=(4,0))
        ttk.Label(g2, textvariable=self.stat_wh_keys, bootstyle=SECONDARY).grid(row=1, column=1, sticky="e", pady=(4,0))

        ttk.Separator(lf).pack(fill=X, pady=8)

        # Last diff + output dir
        g3 = ttk.Frame(lf); g3.pack(fill=X)
        ttk.Label(g3, text="Last diff count").grid(row=0, column=0, sticky="w")
        ttk.Label(g3, textvariable=self.stat_last_diff, bootstyle=WARNING).grid(row=0, column=1, sticky="e")
        ttk.Label(g3, text="Output folder").grid(row=1, column=0, sticky="w", pady=(4,0))
        ttk.Label(g3, textvariable=self.stat_output_dir, bootstyle=SECONDARY, wraplength=240).grid(row=1, column=1, sticky="e", pady=(4,0))

        ttk.Separator(lf).pack(fill=X, pady=8)

        # Automation state
        g4 = ttk.Frame(lf); g4.pack(fill=X)
        ttk.Label(g4, text="Automation").grid(row=0, column=0, sticky="w")
        ttk.Label(g4, textvariable=self.stat_auto_state, bootstyle=INFO).grid(row=0, column=1, sticky="e")
        ttk.Label(g4, text="Last run").grid(row=1, column=0, sticky="w", pady=(4,0))
        ttk.Label(g4, textvariable=self.stat_last_run, bootstyle=SECONDARY).grid(row=1, column=1, sticky="e", pady=(4,0))
        ttk.Label(g4, text="Next run").grid(row=2, column=0, sticky="w")
        ttk.Label(g4, textvariable=self.stat_next_run, bootstyle=SECONDARY).grid(row=2, column=1, sticky="e")

    def _build_export_block(self, parent: ttk.Frame):
        lf = ttk.LabelFrame(parent, text="Export", padding=10)
        lf.pack(fill=X, pady=(0, 4))

        row1 = ttk.Frame(lf); row1.pack(fill=X)
        ttk.Label(row1, text="Output Folder").pack(side=LEFT)
        e = ttk.Entry(row1, textvariable=self.var_out_dir); e.pack(side=LEFT, fill=X, expand=True, padx=6)
        ttk.Button(row1, text="Choose", command=self._pick_out_dir).pack(side=LEFT)

        row2 = ttk.Frame(lf); row2.pack(fill=X, pady=(6, 0))
        ttk.Label(row2, text="Chunk size").pack(side=LEFT)
        ttk.Entry(row2, width=10, textvariable=self.var_chunk).pack(side=LEFT, padx=6)
        ttk.Button(lf, text="Generate Magento 2 CSV", bootstyle=SUCCESS, command=self._generate).pack(fill=X, pady=(10, 0))

    def _build_main_area(self, parent: ttk.Frame):
        self.nb = ttk.Notebook(parent)
        self.nb.grid(row=0, column=0, sticky="ew", pady=(0,6))

        # Tab: Mapping Overview
        tbar = ttk.Frame(self.nb, padding=6)
        self.nb.add(tbar, text="Mapping Overview")
        ttk.Label(tbar, text="Preview:").pack(side=LEFT)
        cb = ttk.Combobox(tbar, state="readonly", values=["ECommerce", "Warehouse"], textvariable=self.current_preview_side, width=12)
        cb.pack(side=LEFT, padx=(6,10))
        ttk.Button(tbar, text="Show 50 rows", command=self._preview_rows).pack(side=LEFT)

        # NEW Tab: Composite Keys (moved here)
        self.tab_keys = ttk.Frame(self.nb, padding=10)
        self.nb.add(self.tab_keys, text="Composite Keys")
        self.tab_keys.columnconfigure(0, weight=1)
        self.tab_keys.columnconfigure(1, weight=1)

        self._build_composite_side(self.tab_keys, side="ECommerce", col=0)
        self._build_composite_side(self.tab_keys, side="Warehouse", col=1)

        # Tab: Settings & Automation
        self.tab_settings = ttk.Frame(self.nb, padding=10)
        self.nb.add(self.tab_settings, text="Settings & Automation")
        self._build_settings_tab(self.tab_settings)

        # Shared preview table
        pv = ttk.Frame(parent)
        pv.grid(row=1, column=0, sticky="nsew")
        pv.rowconfigure(1, weight=1); pv.columnconfigure(0, weight=1)

        self.preview_tree = ttk.Treeview(pv, show="headings", bootstyle="info")
        self.preview_tree.grid(row=1, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(pv, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=1, column=1, sticky="ns")

        # Logs
        lf = ttk.LabelFrame(parent, text="Logs", padding=6)
        lf.grid(row=2, column=0, sticky="ew")
        self.log_text = tk.Text(lf, height=7, wrap="word")
        self.log_text.pack(fill=X)

    # ------- Composite Keys Tab (per side) -------
    def _build_composite_side(self, parent: ttk.Frame, side: str, col: int):
        lf = ttk.LabelFrame(parent, text=f"{side} Mapping", padding=10)
        lf.grid(row=0, column=col, sticky="nsew", padx=(0,8) if col == 0 else (8,0))
        lf.columnconfigure(0, weight=1)

        # column list
        ttk.Label(lf, text="Available columns").grid(row=0, column=0, sticky="w")
        fr = ttk.Frame(lf); fr.grid(row=1, column=0, sticky="nsew")
        lb = tk.Listbox(fr, selectmode="extended", height=10, exportselection=False)
        lb.pack(side=LEFT, fill=BOTH, expand=True)
        sb = ttk.Scrollbar(fr, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set); sb.pack(side=LEFT, fill=Y, padx=(4, 0))

        # key builder
        kb = ttk.LabelFrame(lf, text="Composite Key (ordered)", padding=8)
        kb.grid(row=2, column=0, sticky="nsew", pady=(8,0))
        lb_key = tk.Listbox(kb, selectmode="browse", height=6, exportselection=False)
        lb_key.pack(fill=X)
        btns = ttk.Frame(kb); btns.pack(fill=X, pady=4)
        ttk.Button(btns, text="Add →", command=lambda s=side: self._key_add(s)).pack(side=LEFT)
        ttk.Button(btns, text="← Remove", command=lambda s=side: self._key_remove(s)).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="↑ Up", command=lambda s=side: self._key_up(s)).pack(side=LEFT, padx=(12, 6))
        ttk.Button(btns, text="↓ Down", command=lambda s=side: self._key_down(s)).pack(side=LEFT)

        # qty + sku dropdowns
        row = ttk.Frame(lf); row.grid(row=3, column=0, sticky="ew", pady=(8,0))
        ttk.Label(row, text="Qty column").pack(side=LEFT)
        cb_qty = ttk.Combobox(row, state="readonly"); cb_qty.pack(side=LEFT, fill=X, expand=True, padx=6)

        row2 = ttk.Frame(lf); row2.grid(row=4, column=0, sticky="ew", pady=(4,0))
        ttk.Label(row2, text="SKU column").pack(side=LEFT)
        cb_sku = ttk.Combobox(row2, state="readonly"); cb_sku.pack(side=LEFT, fill=X, expand=True, padx=6)

        # normalization
        normf = ttk.LabelFrame(lf, text="Normalization", padding=6)
        normf.grid(row=5, column=0, sticky="ew", pady=(8,0))
        var_trim = tk.BooleanVar(value=True)
        var_case = tk.BooleanVar(value=True)
        var_ws = tk.BooleanVar(value=True)
        var_zpad = tk.IntVar(value=0)

        ttk.Checkbutton(normf, text="Trim", variable=var_trim).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(normf, text="Casefold", variable=var_case).grid(row=0, column=1, sticky="w", padx=(8,0))
        ttk.Checkbutton(normf, text="Collapse spaces", variable=var_ws).grid(row=0, column=2, sticky="w", padx=(8,0))
        ttk.Label(normf, text="Zero-pad width").grid(row=1, column=0, sticky="w", pady=(4,0))
        ttk.Entry(normf, width=6, textvariable=var_zpad).grid(row=1, column=1, sticky="w", pady=(4,0))

        bundle = {
            "lb_cols": lb,
            "lb_key": lb_key,
            "cb_qty": cb_qty,
            "cb_sku": cb_sku,
            "var_trim": var_trim,
            "var_case": var_case,
            "var_ws": var_ws,
            "var_zpad": var_zpad,
        }
        if side == "ECommerce":
            self.widgets_ec = bundle
        else:
            self.widgets_wh = bundle

    # ------- Settings & Automation Tab -------
    def _build_settings_tab(self, parent: ttk.Frame):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        # Source codes editor
        srcf = ttk.LabelFrame(parent, text="Source Codes (exported per row)", padding=10)
        srcf.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        self.lb_sources = tk.Listbox(srcf, height=8, exportselection=False)
        self.lb_sources.pack(side=LEFT, fill=BOTH, expand=True)
        sb = ttk.Scrollbar(srcf, orient="vertical", command=self.lb_sources.yview)
        self.lb_sources.configure(yscrollcommand=sb.set); sb.pack(side=LEFT, fill=Y)
        for sc in self.export_cfg.source_codes:
            self.lb_sources.insert("end", sc)
        btns = ttk.Frame(srcf); btns.pack(side=LEFT, fill=Y, padx=8)
        ttk.Button(btns, text="Add", command=self._source_add).pack(fill=X)
        ttk.Button(btns, text="Remove", command=self._source_remove, bootstyle=WARNING).pack(fill=X, pady=(6,0))
        ttk.Button(btns, text="Reset", command=self._source_reset).pack(fill=X, pady=(6,0))

        # Export defaults
        exf = ttk.LabelFrame(parent, text="Export Defaults", padding=10)
        exf.grid(row=0, column=1, sticky="nsew")
        r1 = ttk.Frame(exf); r1.pack(fill=X, pady=(0,4))
        ttk.Label(r1, text="Output Folder").pack(side=LEFT)
        ttk.Entry(r1, textvariable=self.var_out_dir).pack(side=LEFT, fill=X, expand=True, padx=6)
        ttk.Button(r1, text="Choose", command=self._pick_out_dir).pack(side=LEFT)
        r2 = ttk.Frame(exf); r2.pack(fill=X, pady=(4,0))
        ttk.Label(r2, text="Chunk size").pack(side=LEFT)
        ttk.Entry(r2, width=10, textvariable=self.var_chunk).pack(side=LEFT, padx=6)

        # Automation
        auf = ttk.LabelFrame(parent, text="Automation", padding=10)
        auf.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8,0))
        rowa = ttk.Frame(auf); rowa.pack(fill=X)
        ttk.Checkbutton(rowa, text="Enable automation", variable=self.var_auto_enabled).pack(side=LEFT)
        ttk.Checkbutton(rowa, text="Autorun on start", variable=self.var_auto_autorun).pack(side=LEFT, padx=12)
        rowb = ttk.Frame(auf); rowb.pack(fill=X, pady=(6,0))
        ttk.Label(rowb, text="Poll (seconds)").pack(side=LEFT)
        ttk.Entry(rowb, width=8, textvariable=self.var_auto_seconds).pack(side=LEFT, padx=6)
        ttk.Checkbutton(rowb, text="Run only when files changed", variable=self.var_auto_run_if_changed).pack(side=LEFT, padx=12)

        rowc = ttk.Frame(auf); rowc.pack(fill=X, pady=(8,0))
        ttk.Label(rowc, text="Watch ECommerce CSV").pack(side=LEFT)
        ttk.Entry(rowc, textvariable=self.var_watch_ec).pack(side=LEFT, fill=X, expand=True, padx=6)
        ttk.Button(rowc, text="Browse", command=lambda: self._pick_watch_file(self.var_watch_ec)).pack(side=LEFT)

        rowd = ttk.Frame(auf); rowd.pack(fill=X, pady=(6,0))
        ttk.Label(rowd, text="Watch Warehouse CSV").pack(side=LEFT)
        ttk.Entry(rowd, textvariable=self.var_watch_wh).pack(side=LEFT, fill=X, expand=True, padx=6)
        ttk.Button(rowd, text="Browse", command=lambda: self._pick_watch_file(self.var_watch_wh)).pack(side=LEFT)

        ctl = ttk.Frame(auf); ctl.pack(fill=X, pady=(10,0))
        ttk.Button(ctl, text="Start Automation", bootstyle=SUCCESS, command=self._start_automation).pack(side=LEFT)
        ttk.Button(ctl, text="Stop", bootstyle=WARNING, command=self._stop_automation).pack(side=LEFT, padx=6)

        # Profiles
        prf = ttk.LabelFrame(parent, text="Profile", padding=10)
        prf.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8,0))
        r = ttk.Frame(prf); r.pack(fill=X)
        ttk.Entry(r, textvariable=self.var_profile_path).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(r, text="Browse…", command=self._browse_profile_path).pack(side=LEFT, padx=6)
        ctl2 = ttk.Frame(prf); ctl2.pack(fill=X, pady=(6,0))
        ttk.Button(ctl2, text="Save Profile", bootstyle=SECONDARY, command=self._save_profile).pack(side=LEFT)
        ttk.Button(ctl2, text="Load Profile", command=self._load_profile).pack(side=LEFT, padx=6)

    # ---------------- File & Columns ----------------
    def _pick_ec(self):
        p = filedialog.askopenfilename(title="Select ECommerce CSV", filetypes=[("CSV","*.csv"),("All","*.*")])
        if not p: return
        self.ec_entry.delete(0, "end"); self.ec_entry.insert(0, p)

    def _pick_wh(self):
        p = filedialog.askopenfilename(title="Select Warehouse CSV", filetypes=[("CSV","*.csv"),("All","*.*")])
        if not p: return
        self.wh_entry.delete(0, "end"); self.wh_entry.insert(0, p)

    def _load_columns(self):
        ec_path = self.ec_entry.get().strip()
        wh_path = self.wh_entry.get().strip()
        if not ec_path or not wh_path:
            messagebox.showwarning(APP_NAME, "Please select both CSV files.")
            return

        try:
            ec_cols = read_csv_columns(ec_path)
            wh_cols = read_csv_columns(wh_path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Failed to read columns:\n{e}")
            return

        self.map_ec = Mapping(file_path=ec_path, columns=ec_cols, key_order=[], qty_col=None, sku_col=None)
        self.map_wh = Mapping(file_path=wh_path, columns=wh_cols, key_order=[], qty_col=None, sku_col=None)

        # Fill composite tab widgets
        self._fill_side_widgets("ECommerce", ec_cols)
        self._fill_side_widgets("Warehouse", wh_cols)

        # Stats: row counts + output dir
        self.stat_ec_rows.set(count_csv_rows(ec_path))
        self.stat_wh_rows.set(count_csv_rows(wh_path))
        self.stat_output_dir.set(self.var_out_dir.get())

        # seed automation watch paths if empty
        if not self.var_watch_ec.get():
            self.var_watch_ec.set(ec_path)
        if not self.var_watch_wh.get():
            self.var_watch_wh.set(wh_path)

        self._update_key_stats_from_widgets()
        self._log(f"Loaded columns.\nECommerce: {len(ec_cols)}\nWarehouse: {len(wh_cols)}")

    def _fill_side_widgets(self, side: str, cols: List[str]):
        w = self.widgets_ec if side == "ECommerce" else self.widgets_wh
        w["lb_cols"].delete(0, "end")
        for c in cols:
            w["lb_cols"].insert("end", c)
        w["lb_key"].delete(0, "end")
        w["cb_qty"]["values"] = cols
        w["cb_sku"]["values"] = cols
        guess_qty = next((c for c in cols if c.lower() in ("qty","quantity","freestock","free_stock","stock","onhand")), cols[0] if cols else "")
        guess_sku = next((c for c in cols if "sku" in c.lower()), cols[0] if cols else "")
        w["cb_qty"].set(guess_qty)
        w["cb_sku"].set(guess_sku)
        self._update_key_stats_from_widgets()

    # ---------------- Key Builder Ops ----------------
    def _key_add(self, side: str):
        w = self.widgets_ec if side == "ECommerce" else self.widgets_wh
        lb_cols, lb_key = w["lb_cols"], w["lb_key"]
        for idx in lb_cols.curselection():
            col = lb_cols.get(idx)
            if col not in lb_key.get(0, "end"):
                lb_key.insert("end", col)
        self._update_key_stats_from_widgets()

    def _key_remove(self, side: str):
        w = self.widgets_ec if side == "ECommerce" else self.widgets_wh
        lb_key = w["lb_key"]
        sel = lb_key.curselection()
        if not sel: return
        lb_key.delete(sel[0])
        self._update_key_stats_from_widgets()

    def _key_up(self, side: str):
        w = self.widgets_ec if side == "ECommerce" else self.widgets_wh
        lb_key = w["lb_key"]
        sel = lb_key.curselection()
        if not sel or sel[0] == 0: return
        i = sel[0]
        val = lb_key.get(i)
        lb_key.delete(i)
        lb_key.insert(i-1, val)
        lb_key.selection_set(i-1)
        self._update_key_stats_from_widgets()

    def _key_down(self, side: str):
        w = self.widgets_ec if side == "ECommerce" else self.widgets_wh
        lb_key = w["lb_key"]
        sel = lb_key.curselection()
        if not sel: return
        i = sel[0]
        if i >= lb_key.size()-1: return
        val = lb_key.get(i)
        lb_key.delete(i)
        lb_key.insert(i+1, val)
        lb_key.selection_set(i+1)
        self._update_key_stats_from_widgets()

    def _update_key_stats_from_widgets(self):
        ec_keys = list(self.widgets_ec["lb_key"].get(0, "end"))
        wh_keys = list(self.widgets_wh["lb_key"].get(0, "end"))
        self.stat_ec_keys.set(", ".join(ec_keys) if ec_keys else "not set")
        self.stat_wh_keys.set(", ".join(wh_keys) if wh_keys else "not set")

    # ---------------- Preview ----------------
    def _preview_rows(self):
        side = self.current_preview_side.get()
        mapping = self._collect_mapping(side)
        if not mapping.file_path:
            messagebox.showwarning(APP_NAME, f"Select {side} CSV and load columns first.")
            return
        try:
            cols = mapping.columns
            rows = load_rows(mapping.file_path, cols)[:50]
            self._fill_preview(cols, rows)
            self._log(f"Previewed {side}: {len(rows)} rows")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Preview failed:\n{e}")

    def _fill_preview(self, cols: List[str], rows: List[Dict]):
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = cols
        self.preview_tree["show"] = "headings"
        for c in cols:
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=max(90, min(240, len(c)*10)), anchor="w")
        for r in rows:
            self.preview_tree.insert("", "end", values=[r.get(c, "") for c in cols])

    # ---------------- Export ----------------
    def _pick_out_dir(self):
        d = filedialog.askdirectory(title="Select output folder", initialdir=self.var_out_dir.get())
        if d:
            self.var_out_dir.set(d)
            self.stat_output_dir.set(d)

    def _collect_mapping(self, side: str) -> Mapping:
        is_ec = (side == "ECommerce")
        mapping = self.map_ec if is_ec else self.map_wh
        w = self.widgets_ec if is_ec else self.widgets_wh
        mapping.key_order = list(w["lb_key"].get(0, "end"))
        mapping.qty_col = w["cb_qty"].get() or None
        mapping.sku_col = w["cb_sku"].get() or None
        mapping.normalization = Normalization(
            trim=bool(w["var_trim"].get()),
            casefold=bool(w["var_case"].get()),
            collapse_ws=bool(w["var_ws"].get()),
            zero_pad_width=int(w["var_zpad"].get() or 0),
        )
        return mapping

    def _open_new_chunk(self, out_dir: Path, date_tag: str, counter: int):
        out_path = out_dir / f"m2_stock_import_{date_tag}_{counter}.csv"
        fp = open(out_path, "w", newline="", encoding="utf-8")
        writer = csv.writer(fp)
        writer.writerow(["sku", "stock_status", "source_code", "qty"])
        self._log(f"Opened chunk: {out_path}")
        return writer, fp

    def _generate(self, silent: bool = False):
        self._collect_mapping("ECommerce")
        self._collect_mapping("Warehouse")

        ec = self.map_ec
        wh = self.map_wh

        for side, m in (("ECommerce", ec), ("Warehouse", wh)):
            if not m.file_path:
                if not silent:
                    messagebox.showwarning(APP_NAME, f"{side}: file not set.")
                return
            if not m.key_order:
                if not silent:
                    messagebox.showwarning(APP_NAME, f"{side}: choose composite key columns (Composite Keys tab).")
                return
            if not m.qty_col or not m.sku_col:
                if not silent:
                    messagebox.showwarning(APP_NAME, f"{side}: select Qty & SKU (Composite Keys tab).")
                return

        out_dir = Path(self.var_out_dir.get().strip() or self.export_cfg.out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        chunk = max(100, int(self.var_chunk.get() or 10000))

        try:
            ec_use = sorted(set(ec.key_order + [ec.qty_col, ec.sku_col]))
            wh_use = sorted(set(wh.key_order + [wh.qty_col, wh.sku_col]))
            ec_rows = load_rows(ec.file_path, ec_use)
            wh_rows = load_rows(wh.file_path, wh_use)
            self._log(f"Loaded ECommerce rows: {len(ec_rows)}; Warehouse rows: {len(wh_rows)}")

            ec_index: Dict[str, Tuple[int, str]] = {}
            for r in ec_rows:
                key = build_composite(r, ec.key_order, ec.normalization)
                ec_index[key] = (to_int_safe(r.get(ec.qty_col, "0")), str(r.get(ec.sku_col, "")).strip())

            wh_index: Dict[str, int] = {}
            for r in wh_rows:
                key = build_composite(r, wh.key_order, wh.normalization)
                wh_index[key] = to_int_safe(r.get(wh.qty_col, "0"))

            keys = set(ec_index.keys()) | set(wh_index.keys())
            changes: List[Tuple[str, int, int, str]] = []

            for k in keys:
                ec_qty, sku = ec_index.get(k, (0, ""))    # treat missing as 0
                wh_qty = wh_index.get(k, 0)
                if ec_qty != wh_qty:
                    if not sku:
                        sku = k
                    changes.append((sku, ec_qty, wh_qty, k))

            # Update sidebar stats
            self.stat_last_diff.set(len(changes))
            self.stat_last_run.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

            if not changes:
                self._log("No differences detected.")
                if not silent:
                    messagebox.showinfo(APP_NAME, "No differences detected.")
                return

            date_tag = datetime.now().strftime("%Y%m%d")
            file_counter = 1
            row_counter = 0
            writer, fp = self._open_new_chunk(out_dir, date_tag, file_counter)
            total_rows = 0

            source_codes = self._current_source_codes()

            for sku, new_qty, _old_qty, _key in changes:
                stock_status = 1 if new_qty > 0 else 0
                for sc in source_codes:
                    writer.writerow([sku, stock_status, sc, new_qty])
                    row_counter += 1
                    total_rows += 1
                    if row_counter >= chunk:
                        fp.close()
                        file_counter += 1
                        row_counter = 0
                        writer, fp = self._open_new_chunk(out_dir, date_tag, file_counter)

            fp.close()
            self._log(f"Generated {file_counter} file(s); {total_rows} rows. Output: {out_dir}")
            if not silent:
                messagebox.showinfo(APP_NAME, f"Done.\nFiles written to:\n{out_dir}")

            # refresh watch mtimes after successful run
            self._refresh_watch_mtimes()

        except Exception as e:
            self._log("ERROR:\n" + traceback.format_exc(), style="danger")
            if not silent:
                messagebox.showerror(APP_NAME, f"Export failed:\n{e}")

    # ---------------- Settings/Profiles/Automation helpers ----------------
    def _current_source_codes(self) -> List[str]:
        return [self.lb_sources.get(i) for i in range(self.lb_sources.size())] or self.export_cfg.source_codes

    def _source_add(self):
        val = tk.simpledialog.askstring("Add source_code", "Enter source_code value:")
        if val:
            self.lb_sources.insert("end", val.strip())

    def _source_remove(self):
        sel = self.lb_sources.curselection()
        if not sel: return
        for i in reversed(sel):
            self.lb_sources.delete(i)

    def _source_reset(self):
        self.lb_sources.delete(0, "end")
        for sc in ["pos_337", "src_virtualstock"]:
            self.lb_sources.insert("end", sc)

    def _browse_profile_path(self):
        p = filedialog.asksaveasfilename(
            title="Save profile as",
            defaultextension=PROFILE_EXT,
            filetypes=[("M2 Profile", f"*{PROFILE_EXT}"), ("JSON", "*.json"), ("All", "*.*")]
        )
        if p: self.var_profile_path.set(p)

    def _save_profile(self):
        path = self.var_profile_path.get().strip() or filedialog.asksaveasfilename(
            title="Save profile as",
            defaultextension=PROFILE_EXT,
            filetypes=[("M2 Profile", f"*{PROFILE_EXT}")]
        )
        if not path: return

        # collect latest mapping/opts
        self._collect_mapping("ECommerce")
        self._collect_mapping("Warehouse")
        data = {
            "export_cfg": {
                "out_dir": self.var_out_dir.get(),
                "chunk_size": int(self.var_chunk.get() or 10000),
                "source_codes": self._current_source_codes(),
            },
            "mapping_ec": self._mapping_to_dict(self.map_ec),
            "mapping_wh": self._mapping_to_dict(self.map_wh),
            "automation": {
                "enabled": bool(self.var_auto_enabled.get()),
                "autorun_on_start": bool(self.var_auto_autorun.get()),
                "poll_seconds": int(self.var_auto_seconds.get() or 30),
                "watch_ec_path": self.var_watch_ec.get(),
                "watch_wh_path": self.var_watch_wh.get(),
                "run_if_changed": bool(self.var_auto_run_if_changed.get()),
            },
        }
        Path(path).write_text(json.dumps(data, indent=2), encoding="utf-8")
        self._log(f"Profile saved: {path}", "success")
        self.var_profile_path.set(path)

    def _load_profile(self):
        path = self.var_profile_path.get().strip() or filedialog.askopenfilename(
            title="Open profile",
            filetypes=[("M2 Profile", f"*{PROFILE_EXT}"), ("JSON", "*.json"), ("All", "*.*")]
        )
        if not path: return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))

            # export cfg
            ecfg = data.get("export_cfg", {})
            self.var_out_dir.set(ecfg.get("out_dir", self.var_out_dir.get()))
            self.var_chunk.set(int(ecfg.get("chunk_size", self.var_chunk.get())))
            self.lb_sources.delete(0, "end")
            for sc in ecfg.get("source_codes", self.export_cfg.source_codes):
                self.lb_sources.insert("end", sc)
            self.stat_output_dir.set(self.var_out_dir.get())

            # mappings
            self.map_ec = self._dict_to_mapping(data.get("mapping_ec", {}))
            self.map_wh = self._dict_to_mapping(data.get("mapping_wh", {}))

            # widgets in Composite Keys tab
            if self.map_ec.columns:
                self._fill_side_widgets("ECommerce", self.map_ec.columns)
                self._apply_mapping_to_widgets("ECommerce", self.map_ec)
            if self.map_wh.columns:
                self._fill_side_widgets("Warehouse", self.map_wh.columns)
                self._apply_mapping_to_widgets("Warehouse", self.map_wh)

            # automation
            aut = data.get("automation", {})
            self.var_auto_enabled.set(bool(aut.get("enabled", self.var_auto_enabled.get())))
            self.var_auto_autorun.set(bool(aut.get("autorun_on_start", self.var_auto_autorun.get())))
            self.var_auto_seconds.set(int(aut.get("poll_seconds", self.var_auto_seconds.get())))
            self.var_watch_ec.set(aut.get("watch_ec_path", self.var_watch_ec.get()))
            self.var_watch_wh.set(aut.get("watch_wh_path", self.var_watch_wh.get()))
            self.var_auto_run_if_changed.set(bool(aut.get("run_if_changed", self.var_auto_run_if_changed.get())))

            # stats
            if self.map_ec.file_path: self.stat_ec_rows.set(count_csv_rows(self.map_ec.file_path))
            if self.map_wh.file_path: self.stat_wh_rows.set(count_csv_rows(self.map_wh.file_path))
            self._update_key_stats_from_widgets()

            self.var_profile_path.set(path)
            self._log(f"Profile loaded: {path}", "success")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Failed to load profile:\n{e}")

    def _mapping_to_dict(self, m: Mapping) -> Dict:
        return {
            "file_path": m.file_path,
            "columns": m.columns or [],
            "key_order": m.key_order or [],
            "qty_col": m.qty_col,
            "sku_col": m.sku_col,
            "normalization": asdict(m.normalization),
        }

    def _dict_to_mapping(self, d: Dict) -> Mapping:
        norm = d.get("normalization", {})
        return Mapping(
            file_path=d.get("file_path"),
            columns=d.get("columns") or [],
            key_order=d.get("key_order") or [],
            qty_col=d.get("qty_col"),
            sku_col=d.get("sku_col"),
            normalization=Normalization(
                trim=bool(norm.get("trim", True)),
                casefold=bool(norm.get("casefold", True)),
                collapse_ws=bool(norm.get("collapse_ws", True)),
                zero_pad_width=int(norm.get("zero_pad_width", 0)),
            )
        )

    def _apply_mapping_to_widgets(self, side: str, m: Mapping):
        w = self.widgets_ec if side == "ECommerce" else self.widgets_wh
        w["lb_key"].delete(0, "end")
        for c in m.key_order:
            w["lb_key"].insert("end", c)
        w["cb_qty"].set(m.qty_col or "")
        w["cb_sku"].set(m.sku_col or "")
        w["var_trim"].set(m.normalization.trim)
        w["var_case"].set(m.normalization.casefold)
        w["var_ws"].set(m.normalization.collapse_ws)
        w["var_zpad"].set(m.normalization.zero_pad_width)
        self._update_key_stats_from_widgets()

    # ---------------- Automation ----------------
    def _pick_watch_file(self, var: tk.StringVar):
        p = filedialog.askopenfilename(title="Watch CSV", filetypes=[("CSV","*.csv"),("All","*.*")])
        if p:
            var.set(p)

    def _start_automation(self):
        if self._auto_running:
            self._log("Automation already running.", "warning")
            return
        if not self.var_auto_enabled.get():
            self._log("Enable automation first.", "warning")
            return
        self._auto_running = True
        self.stat_auto_state.set("Running")
        self._log("Automation started.", "info")
        self._refresh_watch_mtimes()
        self._schedule_auto_tick()

    def _stop_automation(self):
        if not self._auto_running:
            return
        self._auto_running = False
        if self._auto_job_id:
            try:
                self.root.after_cancel(self._auto_job_id)
            except Exception:
                pass
            self._auto_job_id = None
        self._next_tick_eta = None
        self.stat_next_run.set("–")
        self.stat_auto_state.set("Stopped")
        self._log("Automation stopped.", "warning")

    def _schedule_auto_tick(self):
        if not self._auto_running:
            return
        secs = max(5, int(self.var_auto_seconds.get() or 30))
        self._next_tick_eta = datetime.now() + timedelta(seconds=secs)
        self.stat_next_run.set(self._next_tick_eta.strftime("%Y-%m-%d %H:%M:%S"))
        self._auto_job_id = self.root.after(secs * 1000, self._auto_tick)

    def _refresh_watch_mtimes(self):
        ec = self.var_watch_ec.get().strip()
        wh = self.var_watch_wh.get().strip()
        try:
            self.auto_cfg._last_mtime_ec = Path(ec).stat().st_mtime if ec and Path(ec).exists() else 0.0
        except Exception:
            self.auto_cfg._last_mtime_ec = 0.0
        try:
            self.auto_cfg._last_mtime_wh = Path(wh).stat().st_mtime if wh and Path(wh).exists() else 0.0
        except Exception:
            self.auto_cfg._last_mtime_wh = 0.0

    def _auto_tick(self):
        if not self._auto_running:
            return
        try:
            ec = self.var_watch_ec.get().strip()
            wh = self.var_watch_wh.get().strip()
            ec_m = Path(ec).stat().st_mtime if ec and Path(ec).exists() else 0.0
            wh_m = Path(wh).stat().st_mtime if wh and Path(wh).exists() else 0.0

            changed = (ec_m != self.auto_cfg._last_mtime_ec) or (wh_m != self.auto_cfg._last_mtime_wh)
            should_run = True if not self.var_auto_run_if_changed.get() else changed

            if should_run:
                self._log("Automation: generating (triggered).", "info")
                if ec and not self.map_ec.file_path:
                    self.map_ec.file_path = ec
                    self.map_ec.columns = read_csv_columns(ec)
                    self._fill_side_widgets("ECommerce", self.map_ec.columns)
                if wh and not self.map_wh.file_path:
                    self.map_wh.file_path = wh
                    self.map_wh.columns = read_csv_columns(wh)
                    self._fill_side_widgets("Warehouse", self.map_wh.columns)

                self._generate(silent=True)
                self._refresh_watch_mtimes()
            else:
                self._log("Automation: no change detected.", "secondary")
        except Exception:
            self._log("Automation error:\n" + traceback.format_exc(), "danger")
        finally:
            self._schedule_auto_tick()

    # ---------------- App Utils ----------------
    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self._preview_rows())
        self.root.bind("<Control-q>", lambda e: self.root.destroy())
        self.root.bind("<F5>", lambda e: self._generate())

    def _log(self, text: str, style: str = "secondary"):
        self.status.config(text=text, bootstyle=style)
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
        self.log_text.see("end")

# -------------------------- Entrypoint -----------------------------
def main():
    win = Window(themename=DEFAULT_THEME)
    # win.style = Style(theme=DEFAULT_THEME)
    app = CompositeKeyDashboard(win)
    win.mainloop()

if __name__ == "__main__":
    main()

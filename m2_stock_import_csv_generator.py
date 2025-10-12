import os
import sys
import json
import csv
import time
import threading
import queue
from pathlib import Path
from typing import Optional
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
SIDEBAR_WIDTH = 280
RIGHT_MIN_WIDTH = 400
GROUP_PAD = 10
GRID_PAD = 8
BOOTSTYLE_PROGRESS = "warning"
PREVIEW_ROWS = 100
DEFAULT_SOURCE_CODES = ["pos_337", "src_virtualstock"]
APP_HOME = Path.home() / ".m2_stock_app"
APP_HOME.mkdir(exist_ok=True)
PROFILE_FILE = APP_HOME / "profile.json"
def sniff_csv_meta(file_path: str) -> tuple[str, str]:
    enc_candidates = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    with open(file_path, "rb") as f:
        head = f.read(2048)
    dialect = csv.Sniffer().sniff(head.decode("utf-8", errors="ignore"), delimiters=",;\t|")
    for enc in enc_candidates:
        try:
            with open(file_path, "r", encoding=enc) as fh:
                fh.read(1024)
            return enc, dialect.delimiter
        except Exception:
            continue
    return "utf-8", dialect.delimiter
def coerce_qty_series(series: pd.Series, rule: str, clamp_negative: bool) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    rule = (rule or "round").lower()
    if rule == "floor":
        s = s.apply(lambda x: int(x // 1))
    elif rule == "ceil":
        s = s.apply(lambda x: int(x) if x == int(x) else int(x) + 1)
    else:
        s = s.round(0).astype(int)
    if clamp_negative:
        s = s.clip(lower=0)
    return s.astype(int)
class M2StockApp(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("M2 Stock Import CSV Generator")
        self.geometry("750x500")
        self.m2_df: Optional[pd.DataFrame] = None
        self.preview_base_df: Optional[pd.DataFrame] = None
        self.original_file_path: Optional[str] = None
        self.source_codes: list[str] = DEFAULT_SOURCE_CODES.copy()
        self.output_folder = os.path.expanduser("~/Downloads")
        self.chunk_size = 1000
        self.available_columns: list[str] = []
        self.sku_column = tk.StringVar(value="")
        self.qty_column = tk.StringVar(value="")
        self.use_raw_sku = tk.BooleanVar(value=False)
        self.qty_rule = tk.StringVar(value="round")
        self.clamp_negative = tk.BooleanVar(value=True)
        self.combine_single_file = tk.BooleanVar(value=False)
        self.theme_var = tk.StringVar(value="darkly")
        self.search_var = tk.StringVar(value="")
        self._worker_q: queue.Queue = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._busy = tk.BooleanVar(value=False)
        self._load_started_at: Optional[float] = None
        self.build_ui()
        self.bind_shortcuts()
        self.try_load_profile()
    def build_ui(self):
        root_pane = tb.Panedwindow(self, orient="horizontal")
        root_pane.pack(fill=BOTH, expand=True)
        sidebar = tb.Frame(root_pane, padding=(12, 12))
        sidebar.configure(width=SIDEBAR_WIDTH)
        self.stats_label = tb.Label(sidebar, justify=LEFT, anchor=NW)
        self.stats_label.pack(anchor=NW, pady=(0, GRID_PAD))
        self._render_stats(None)  # pretty initial stats
        root_pane.add(sidebar)
        right = tb.Frame(root_pane)
        root_pane.add(right)
        root_pane.pane(sidebar, weight=0)
        root_pane.pane(right,   weight=1)
        self.after_idle(lambda: root_pane.sashpos(0, SIDEBAR_WIDTH))
        self.notebook = tb.Notebook(right)
        self.notebook.pack(fill=BOTH, expand=True)
        cfg = tb.Frame(self.notebook, padding=12)
        self.notebook.add(cfg, text="Configuration")
        cfg.columnconfigure(0, weight=1)
        r = 0
        tb.Label(cfg, text="1) Browse file (CSV or Excel):").grid(row=r, column=0, sticky=W, pady=(0, 4)); r += 1
        self.entry_file_path = tb.Entry(cfg, state="readonly")
        self.entry_file_path.grid(row=r, column=0, sticky=EW); r += 1
        ToolTip(self.entry_file_path, text="Selected file path (read-only).")
        row_browse = tb.Frame(cfg); row_browse.grid(row=r, column=0, sticky=EW, pady=(6, 8)); r += 1
        btn_browse = tb.Button(row_browse, text="Browse", command=self.select_file)
        btn_browse.pack(side=LEFT)
        ToolTip(btn_browse, text="Pick a CSV/XLSX/XLS file. The app will auto-load file structure.")
        tb.Label(cfg, text="2) Load file structure (auto after Browse)").grid(row=r, column=0, sticky=W, pady=(4, 12)); r += 1
        tb.Label(cfg, text="3) Choose SKU column:").grid(row=r, column=0, sticky=W, pady=(0, 4)); r += 1
        self.dropdown_sku = tb.Combobox(cfg, textvariable=self.sku_column, state="disabled")
        self.dropdown_sku.grid(row=r, column=0, sticky=EW); r += 1
        self.dropdown_sku.bind("<<ComboboxSelected>>", self._workflow_update)
        ToolTip(self.dropdown_sku, text="Select the SKU column from the file.")
        tb.Label(cfg, text="3) Choose Quantity column:").grid(row=r, column=0, sticky=W, pady=(10, 4)); r += 1
        self.dropdown_qty = tb.Combobox(cfg, textvariable=self.qty_column, state="disabled")
        self.dropdown_qty.grid(row=r, column=0, sticky=EW); r += 1
        self.dropdown_qty.bind("<<ComboboxSelected>>", self._workflow_update)
        ToolTip(self.dropdown_qty, text="Select the Quantity column from the file.")
        tb.Label(cfg, text="4) Unlock Preview:").grid(row=r, column=0, sticky=W, pady=(12, 4)); r += 1
        self.btn_unlock = tb.Button(cfg, text="Unlock Preview", command=self.unlock_preview, state=DISABLED)
        self.btn_unlock.grid(row=r, column=0, sticky=W); r += 1
        ToolTip(self.btn_unlock, text="Runs the transform with your selected columns and unlocks the Preview tab.")
        st = tb.Frame(self.notebook, padding=12)
        self.notebook.add(st, text="Settings")
        st.columnconfigure(0, weight=1, uniform="stcols")
        st.columnconfigure(1, weight=1, uniform="stcols")
        source_frame = tb.LabelFrame(st, text="Source Codes", padding=GROUP_PAD)
        source_frame.grid(row=0, column=0, sticky=EW, padx=(0, GRID_PAD))
        source_frame.columnconfigure(0, weight=1)
        self.listbox_sources = tk.Listbox(source_frame, height=7)
        self.listbox_sources.grid(row=0, column=0, sticky=EW)
        sbtn = tb.Frame(source_frame); sbtn.grid(row=1, column=0, sticky=W, pady=(6, 0))
        b_add = tb.Button(sbtn, text="Add", command=self.add_source_code); b_add.pack(side=LEFT, padx=(0, 4))
        b_remove = tb.Button(sbtn, text="Remove", command=self.remove_source_code); b_remove.pack(side=LEFT, padx=4)
        b_reset = tb.Button(sbtn, text="Reset", command=self.reset_source_codes); b_reset.pack(side=LEFT, padx=4)
        ToolTip(b_add, "Add a source_code value."); ToolTip(b_remove, "Remove selected source_code."); ToolTip(b_reset, "Reset to defaults.")
        self.refresh_source_list()
        rules = tb.LabelFrame(st, text="Quantity Handling", padding=GROUP_PAD)
        rules.grid(row=0, column=1, sticky=NW, padx=(GRID_PAD, 0))
        rules.columnconfigure(0, weight=1)
        tb.Label(rules, text="Whole-number rule:").grid(row=0, column=0, sticky=W, pady=(0, 4))
        cb_rule = tb.Combobox(rules, textvariable=self.qty_rule, state="readonly", values=["round", "floor", "ceil"])
        cb_rule.grid(row=1, column=0, sticky=W)
        chk_clamp = tb.Checkbutton(rules, text="Clamp negatives to 0", variable=self.clamp_negative)
        chk_clamp.grid(row=2, column=0, sticky=W, pady=(10, 0))
        ToolTip(cb_rule, "Round, floor, or ceil when coercing quantities."); ToolTip(chk_clamp, "If enabled, negatives become 0.")
        exp = tb.LabelFrame(st, text="Export Settings", padding=GROUP_PAD)
        exp.grid(row=1, column=0, sticky=EW, padx=(0, GRID_PAD), pady=(GRID_PAD, 0))
        exp.columnconfigure(0, weight=1)
        tb.Label(exp, text="Output Folder:").grid(row=0, column=0, sticky=W, pady=(0, 4))
        out_row = tb.Frame(exp); out_row.grid(row=1, column=0, sticky=EW); out_row.columnconfigure(0, weight=1)
        self.entry_output_folder = tb.Entry(out_row)
        self.entry_output_folder.insert(0, self.output_folder)
        self.entry_output_folder.grid(row=0, column=0, sticky=EW, padx=(0, 6))
        b_choose = tb.Button(out_row, text="Choose", command=self.choose_output_folder); b_choose.grid(row=0, column=1, sticky=E)
        ToolTip(self.entry_output_folder, "Where to write exported CSV files."); ToolTip(b_choose, "Select an output directory.")
        tb.Label(exp, text="Chunk Size:").grid(row=2, column=0, sticky=W, pady=(10, 4))
        vcmd = (self.register(self._validate_int), "%P")
        self.entry_chunk_size = tb.Entry(exp, width=12, validate="key", validatecommand=vcmd)
        self.entry_chunk_size.insert(0, str(self.chunk_size))
        self.entry_chunk_size.grid(row=3, column=0, sticky=W)
        chk_single = tb.Checkbutton(exp, text="Export single file (ignore chunking)", variable=self.combine_single_file)
        chk_single.grid(row=4, column=0, sticky=W, pady=(10, 0))
        ToolTip(self.entry_chunk_size, "Rows per part file (ignored if single-file)."); ToolTip(chk_single, "Write one combined CSV instead of parts.")
        app_box = tb.LabelFrame(st, text="Appearance & Profile", padding=GROUP_PAD)
        app_box.grid(row=1, column=1, sticky=NW, padx=(GRID_PAD, 0), pady=(GRID_PAD, 0))
        app_box.columnconfigure(0, weight=1)
        tb.Label(app_box, text="Theme:").grid(row=0, column=0, sticky=W, pady=(0, 4))
        themes = sorted(tb.Style().theme_names())
        self.combo_theme = tb.Combobox(app_box, state="readonly", values=themes, textvariable=self.theme_var)
        self.combo_theme.grid(row=1, column=0, sticky=W)
        self.combo_theme.bind("<<ComboboxSelected>>", self._on_theme_change)
        ToolTip(self.combo_theme, "Switch between ttkbootstrap themes.")
        tb.Checkbutton(app_box, text="Use 'key' as SKU (no split)", variable=self.use_raw_sku).grid(row=2, column=0, sticky=W, pady=(12, 0))
        prof_row = tb.Frame(app_box); prof_row.grid(row=3, column=0, sticky=W, pady=(12, 0))
        b_save = tb.Button(prof_row, text="Save Profile", command=self.save_profile); b_save.pack(side=LEFT)
        b_load = tb.Button(prof_row, text="Load Profile", command=self.try_load_profile); b_load.pack(side=LEFT, padx=(6, 0))
        ToolTip(b_save, "Save current configuration."); ToolTip(b_load, "Load configuration from profile.")
        tprev = tb.Frame(self.notebook, padding=10)
        self.notebook.add(tprev, text="Preview")
        top_prev = tb.Frame(tprev)
        top_prev.pack(fill=X, pady=(0, 8))
        self.btn_export = tb.Button(top_prev, text="Export M2 CSV", bootstyle=SUCCESS, command=self.export_csv, state=DISABLED)
        self.btn_export.pack(side=LEFT)
        ToolTip(self.btn_export, text="Export transformed data to CSV files.")
        b_open = tb.Button(top_prev, text="Open Output Folder", command=self.open_output_folder)
        b_open.pack(side=LEFT, padx=(6, 0))
        ToolTip(b_open, text="Open the output directory.")
        tb.Frame(top_prev).pack(side=LEFT, expand=True, fill=X)
        tb.Label(top_prev, text="Search:").pack(side=LEFT, padx=(0, 6))
        self.entry_search = tb.Entry(top_prev, textvariable=self.search_var, width=32)
        self.entry_search.pack(side=LEFT)
        ToolTip(self.entry_search, text="Filter preview (any column, case-insensitive).")
        b_clear = tb.Button(top_prev, text="Clear", command=lambda: self.search_var.set(""))
        b_clear.pack(side=LEFT, padx=(6, 0))
        ToolTip(b_clear, "Clear search filter.")
        wrap = tb.Frame(tprev); wrap.pack(fill=BOTH, expand=True)
        self.tree = tb.Treeview(wrap, show="headings")
        vs = tb.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        wrap.columnconfigure(0, weight=1); wrap.rowconfigure(0, weight=1)
        self._bind_mousewheel(self.tree)
        bottom = tb.Frame(self, padding=(8, 4))
        bottom.pack(side=BOTTOM, fill=X)
        self.progress = tb.Progressbar(bottom, mode="indeterminate", bootstyle=BOOTSTYLE_PROGRESS)
        self.progress.pack(fill=X, pady=(0, 4))
        self.status = tb.Label(bottom, text="Ready", anchor=W)
        self.status.pack(fill=X)
        self.search_var.trace_add("write", lambda *_: self._apply_search_filter())
        self._lock_preview_tab(True)
    def bind_shortcuts(self):
        self.bind("<Control-o>", lambda e: self.select_file())
        self.bind("<Control-l>", lambda e: self.load_data())
        self.bind("<Control-e>", lambda e: self.export_csv())
        self.bind("<Control-f>", lambda e: self.entry_search.focus_set())
    def _on_theme_change(self, *_):
        try:
            tb.Style().theme_use(self.theme_var.get())
            self.status.config(text=f"Theme: {self.theme_var.get()}")
        except Exception as e:
            messagebox.showerror("Theme error", str(e))
    def _validate_int(self, val: str) -> bool:
        if val == "":
            return True
        return val.isdigit()
    def _lock_preview_tab(self, lock: bool):
        try:
            self.notebook.tab(2, state="disabled" if lock else "normal")
        except Exception:
            pass
    def _workflow_update(self, *_):
        file_ok = bool(self.entry_file_path.get().strip())
        sku_ok = bool(self.sku_column.get().strip())
        qty_ok = bool(self.qty_column.get().strip())
        if hasattr(self, "btn_unlock"):
            self.btn_unlock.configure(state=NORMAL if (file_ok and sku_ok and qty_ok) else DISABLED)
        if file_ok and sku_ok and qty_ok and self.notebook.tab(2, "state") == "disabled" and not self._busy.get():
            self.after_idle(self.unlock_preview)
    def _set_controls_enabled(self, enabled: bool):
        state = NORMAL if enabled else DISABLED
        try:
            self.btn_load.configure(state=state if self.btn_load['state'] != DISABLED else DISABLED)
        except Exception:
            pass
        try:
            self.btn_export.configure(state=(NORMAL if (enabled and self.m2_df is not None) else DISABLED))
        except Exception:
            pass
    def set_busy(self, busy: bool):
        self._busy.set(busy)
        self._set_controls_enabled(not busy)
        if busy:
            self.progress.start(10)
            self.status.config(text="Working…")
        else:
            self.progress.stop()
    def snapshot_profile(self) -> dict:
        return {
            "file": self.entry_file_path.get().strip(),
            "sku_column": self.sku_column.get().strip(),
            "qty_column": self.qty_column.get().strip(),
            "use_raw_sku": self.use_raw_sku.get(),
            "source_codes": list(self.source_codes),
            "output_folder": self.entry_output_folder.get().strip() or self.output_folder,
            "chunk_size": int(self.entry_chunk_size.get() or "1000"),
            "qty_rule": self.qty_rule.get(),
            "clamp_negative": self.clamp_negative.get(),
            "combine_single_file": self.combine_single_file.get(),
            "theme": self.theme_var.get(),
        }
    def apply_profile(self, prof: dict):
        self.entry_file_path.configure(state="normal")
        self.entry_file_path.delete(0, "end")
        self.entry_file_path.insert(0, prof.get("file", ""))
        self.entry_file_path.configure(state="readonly")
        self.sku_column.set(prof.get("sku_column", ""))
        self.qty_column.set(prof.get("qty_column", ""))
        self.use_raw_sku.set(bool(prof.get("use_raw_sku", False)))
        self.source_codes = list(prof.get("source_codes", DEFAULT_SOURCE_CODES))
        self.refresh_source_list()
        self.output_folder = prof.get("output_folder", self.output_folder)
        self.entry_output_folder.delete(0, "end")
        self.entry_output_folder.insert(0, self.output_folder)
        self.entry_chunk_size.delete(0, "end")
        self.entry_chunk_size.insert(0, str(prof.get("chunk_size", 1000)))
        self.qty_rule.set(prof.get("qty_rule", "round"))
        self.clamp_negative.set(bool(prof.get("clamp_negative", True)))
        self.combine_single_file.set(bool(prof.get("combine_single_file", False)))
        if "theme" in prof:
            self.theme_var.set(prof["theme"])
            self._on_theme_change()
        self._workflow_update()
    def save_profile(self):
        try:
            with open(PROFILE_FILE, "w", encoding="utf-8") as f:
                json.dump(self.snapshot_profile(), f, indent=2)
            messagebox.showinfo("Saved", f"Profile saved to:\n{PROFILE_FILE}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save profile: {e}")
    def try_load_profile(self):
        if PROFILE_FILE.exists():
            try:
                with open(PROFILE_FILE, "r", encoding="utf-8") as f:
                    prof = json.load(f)
                self.apply_profile(prof)
            except Exception:
                pass
    def refresh_source_list(self):
        if not hasattr(self, "listbox_sources"):
            return
        self.listbox_sources.delete(0, "end")
        for code in self.source_codes:
            self.listbox_sources.insert("end", code)
    def add_source_code(self):
        new_code = simpledialog.askstring("Add Source Code", "Enter new source_code:")
        if new_code:
            c = new_code.strip()
            if c and c not in self.source_codes:
                self.source_codes.append(c)
                self.refresh_source_list()
    def remove_source_code(self):
        sel = self.listbox_sources.curselection()
        if sel:
            del self.source_codes[sel[0]]
            self.refresh_source_list()
    def reset_source_codes(self):
        self.source_codes = DEFAULT_SOURCE_CODES.copy()
        self.refresh_source_list()
    def choose_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.entry_output_folder.delete(0, "end")
            self.entry_output_folder.insert(0, folder)
    def open_output_folder(self):
        try:
            os.makedirs(self.output_folder, exist_ok=True)
            if os.name == "nt":
                os.startfile(self.output_folder)
            elif sys.platform == "darwin":
                os.system(f'open "{self.output_folder}"')
            else:
                os.system(f'xdg-open "{self.output_folder}"')
        except Exception:
            messagebox.showinfo("Note", f"Output folder:\n{self.output_folder}")
    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Data files", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return
        self.entry_file_path.configure(state="normal")
        self.entry_file_path.delete(0, "end")
        self.entry_file_path.insert(0, file_path)
        self.entry_file_path.configure(state="readonly")
        self.original_file_path = file_path
        self.load_columns(file_path)
    def load_columns(self, file_path: str):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in (".xlsx", ".xls"):
                try:
                    df = pd.read_excel(file_path, nrows=1)
                except Exception:
                    df = pd.read_excel(file_path, nrows=1, engine="openpyxl")
                enc = "n/a"; delim = "n/a"
            else:
                enc, delim = sniff_csv_meta(file_path)
                df = pd.read_csv(file_path, nrows=1, encoding=enc, sep=delim, on_bad_lines="skip")
            self.available_columns = list(df.columns)
            self.dropdown_sku["values"] = self.available_columns
            self.dropdown_qty["values"] = self.available_columns
            state = "readonly" if self.available_columns else "disabled"
            self.dropdown_sku.configure(state=state)
            self.dropdown_qty.configure(state=state)
            if "key" in self.available_columns:
                self.sku_column.set("key")
            elif "sku" in self.available_columns:
                self.sku_column.set("sku")
            for guess in ("free_stock_tgt", "qty", "quantity", "stock_qty", "free_stock"):
                if guess in self.available_columns:
                    self.qty_column.set(guess)
                    break
            self._workflow_update()
            self.status.config(text=f"Loaded columns ({ext}; enc={enc}, delim='{delim}')")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load columns: {e}")
    def unlock_preview(self):
        if not (self.entry_file_path.get().strip() and self.sku_column.get().strip() and self.qty_column.get().strip()):
            return
        if self._busy.get():
            return
        self._load_started_at = time.perf_counter()
        self.set_busy(True)
        self._worker = threading.Thread(
            target=self._load_process_worker,
            args=(self.entry_file_path.get().strip(), self.snapshot_profile()),
            daemon=True
        )
        self._worker.start()
        self.after(60, self._poll_worker)
    def _poll_worker(self):
        try:
            msg, payload = self._worker_q.get_nowait()
        except queue.Empty:
            if self._worker and self._worker.is_alive():
                self.after(60, self._poll_worker)
                return
            self.set_busy(False)
            return
        if msg == "ok":
            self.m2_df = payload
            self.preview_base_df = self.m2_df.head(PREVIEW_ROWS).copy()
            self.preview_data(self.preview_base_df)
            self.update_stats()
            self._lock_preview_tab(False)
            self.btn_export.configure(state=NORMAL)
            elapsed = 0.0
            if self._load_started_at is not None:
                elapsed = time.perf_counter() - self._load_started_at
            self.set_busy(False)
            self.status.config(text=f"Loaded {len(self.m2_df):,} rows in {elapsed:.2f}s")
        elif msg == "err":
            self.set_busy(False)
            messagebox.showerror("Error", str(payload))
    def _load_process_worker(self, file_path: str, prof: dict):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in (".xlsx", ".xls"):
                try:
                    df = pd.read_excel(file_path)
                except Exception:
                    df = pd.read_excel(file_path, engine="openpyxl")
            else:
                enc, delim = sniff_csv_meta(file_path)
                df = pd.read_csv(file_path, encoding=enc, sep=delim, on_bad_lines="skip")
            sku_col = prof["sku_column"] or ""
            qty_col = prof["qty_column"] or ""
            use_raw = bool(prof.get("use_raw_sku", False))
            qty_rule = prof.get("qty_rule", "round")
            clamp = bool(prof.get("clamp_negative", True))
            sources = prof.get("source_codes", DEFAULT_SOURCE_CODES)
            if not sku_col or not qty_col:
                raise ValueError("SKU and Qty columns are not set.")
            if use_raw:
                df["sku"] = df[sku_col]
            else:
                def split_sku(x):
                    if isinstance(x, str):
                        return x.split("|")[0].strip()
                    return x
                df["sku"] = df[sku_col].apply(split_sku)
            qty_series = coerce_qty_series(df[qty_col], qty_rule, clamp)
            rows = []
            for sku, qty_val in zip(df["sku"], qty_series):
                stock_status = 1 if qty_val > 0 else 0
                for src in sources:
                    rows.append({"sku": sku, "stock_status": stock_status, "source_code": src, "qty": int(qty_val)})
            m2 = pd.DataFrame(rows, columns=["sku", "stock_status", "source_code", "qty"])
            self._worker_q.put(("ok", m2))
        except Exception as e:
            self._worker_q.put(("err", e))
    def preview_data(self, df: pd.DataFrame):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            width = 180 if col == "source_code" else 140
            self.tree.column(col, width=width, anchor=W, stretch=True)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row.values))
    def _apply_search_filter(self):
        if self.preview_base_df is None:
            return
        q = self.search_var.get().strip().lower()
        if not q:
            df = self.preview_base_df
        else:
            mask = self.preview_base_df.apply(
                lambda r: any(q in str(x).lower() for x in r),
                axis=1
            )
            df = self.preview_base_df[mask]
        self.preview_data(df)
    def update_stats(self):
        if self.m2_df is None:
            self._render_stats(None)
            return
        total = int(self.m2_df["sku"].nunique())
        in_stock = int(self.m2_df[self.m2_df["stock_status"] == 1]["sku"].nunique())
        out_stock = int(self.m2_df[self.m2_df["stock_status"] == 0]["sku"].nunique())
        sources = len(self.source_codes)
        stats = {
            "Total SKUs": total,
            "In Stock": in_stock,
            "Out of Stock": out_stock,
            "Source Codes": sources,
        }
        self._render_stats(stats)
    def _render_stats(self, stats: Optional[dict]):
        if not stats:
            text = (
                "📊  Stats\n\n"
                "• Total SKUs: —\n"
                "• In Stock: —\n"
                "• Out of Stock: —\n"
                "• Source Codes: —"
            )
        else:
            text = (
                "📊  Stats\n\n"
                f"• Total SKUs: {stats['Total SKUs']:,}\n"
                f"• In Stock: {stats['In Stock']:,}\n"
                f"• Out of Stock: {stats['Out of Stock']:,}\n"
                f"• Source Codes: {stats['Source Codes']:,}"
            )
        self.stats_label.config(text=text)
    def export_csv(self):
        if self.m2_df is None or self.original_file_path is None:
            messagebox.showwarning("Warning", "No data to export")
            return
        try:
            chunk_size = max(1, int(self.entry_chunk_size.get() or "1000"))
            base_name = os.path.splitext(os.path.basename(self.original_file_path))[0]
            out_dir = self.entry_output_folder.get().strip() or self.output_folder
            os.makedirs(out_dir, exist_ok=True)
            if self.combine_single_file.get():
                output_path = os.path.join(out_dir, f"{base_name}_m2_import.csv")
                tmp = output_path + ".tmp"
                self.m2_df.to_csv(tmp, index=False)
                os.replace(tmp, output_path)
                messagebox.showinfo("Success", f"Exported 1 file to:\n{out_dir}")
                return
            parts = 0
            for i in range(0, len(self.m2_df), chunk_size):
                chunk = self.m2_df.iloc[i: i + chunk_size]
                parts += 1
                output_name = f"{base_name}_m2_import_part{parts}.csv"
                output_path = os.path.join(out_dir, output_name)
                tmp = output_path + ".tmp"
                chunk.to_csv(tmp, index=False)
                os.replace(tmp, output_path)
            messagebox.showinfo("Success", f"Exported {parts} file(s) to:\n{out_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export CSV: {e}")
    def _bind_mousewheel(self, widget: tk.Widget):
        widget.bind("<Enter>", lambda e: widget.focus_set())
        widget.bind("<MouseWheel>", lambda e: self._on_mousewheel(widget, e))
        widget.bind("<Button-4>", lambda e: self._on_mousewheel_linux(widget, -1))
        widget.bind("<Button-5>", lambda e: self._on_mousewheel_linux(widget, 1))
    @staticmethod
    def _on_mousewheel(widget: tk.Widget, event):
        delta = -1 if event.delta > 0 else 1
        try:
            widget.yview_scroll(delta, "units")
        except Exception:
            pass
        return "break"
    @staticmethod
    def _on_mousewheel_linux(widget: tk.Widget, direction: int):
        try:
            widget.yview_scroll(direction, "units")
        except Exception:
            pass
        return "break"
if __name__ == "__main__":
    app = M2StockApp()
    app.mainloop()

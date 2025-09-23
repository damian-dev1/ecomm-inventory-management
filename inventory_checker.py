import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime as dt
import json
from math import floor, ceil
from openpyxl import Workbook
from pathlib import Path

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
            rowheight=20,
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

    def style_text(self, text: tk.Text):
        c = self.colors
        text.configure(bg=c["bg"], fg=c["fg"], insertbackground=c["accent"], selectbackground=c["accent"], selectforeground="#1a1b26", highlightthickness=0, borderwidth=1, relief="flat", wrap="word")

    def tag_palette_for_text(self, text: tk.Text):
        c = self.colors
        text.tag_config("red", foreground=c["error"])
        text.tag_config("orange", foreground=c["warn"])
        text.tag_config("green", foreground=c["success"])
        text.tag_config("blue", foreground=c["info"])
        text.tag_config("bold", font="-weight bold")

    def apply_to_toplevel(self, top: tk.Toplevel):
        top.configure(bg=self.colors["bg"])


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
        self.mode_cb = ttk.Combobox(self.body, textvariable=self.mode_var, state="readonly", values=["strict", "tolerant"], width=12)
        self.mode_cb.grid(row=row, column=1, **pad, sticky="w")
        ttk.Label(self.body, text="Base Method").grid(row=row, column=2, **pad, sticky="e")
        self.base_var = tk.StringVar(value=self.cfg.get("base_method", "max"))
        self.base_cb = ttk.Combobox(self.body, textvariable=self.base_var, state="readonly", values=["max", "avg", "min", "ecommerce", "warehouse"], width=14)
        self.base_cb.grid(row=row, column=3, **pad, sticky="w")
        row += 1
        ttk.Label(self.body, text="Rounding").grid(row=row, column=0, **pad, sticky="e")
        self.round_var = tk.StringVar(value=self.cfg.get("rounding", "floor"))
        self.round_cb = ttk.Combobox(self.body, textvariable=self.round_var, state="readonly", values=["floor", "ceil", "round"], width=10)
        self.round_cb.grid(row=row, column=1, **pad, sticky="w")
        ttk.Label(self.body, text="Auto (equal statuses) →").grid(row=row, column=2, **pad, sticky="e")
        self.auto_equal_var = tk.StringVar(value=self.cfg.get("auto_equal_status_severity", "Medium"))
        self.auto_cb = ttk.Combobox(self.body, textvariable=self.auto_equal_var, state="readonly", values=self.cfg["severity_order"], width=14)
        self.auto_cb.grid(row=row, column=3, **pad, sticky="w")
        row += 1
        ttk.Separator(self.body, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="ew", pady=6)
        row += 1
        ttk.Label(self.body, text="Severity", font="-weight bold").grid(row=row, column=0, **pad)
        ttk.Label(self.body, text="Pct (0.10=10%)", font="-weight bold").grid(row=row, column=1, **pad)
        ttk.Label(self.body, text="Min Units", font="-weight bold").grid(row=row, column=2, **pad)
        ttk.Label(self.body, text="Max Units (blank=None)", font="-weight bold").grid(row=row, column=3, **pad)
        row += 1
        self.entries = {}
        for sev in self.cfg["severity_order"]:
            m = self.cfg["severity_map"][sev]
            ttk.Label(self.body, text=sev).grid(row=row, column=0, **pad, sticky="e")
            v_pct = tk.DoubleVar(value=float(m.get("pct", 0.0)))
            v_min = tk.IntVar(value=int(m.get("min_units", 0) or 0))
            v_max = tk.StringVar(value="" if m.get("max_units") is None else str(int(m["max_units"])))
            ttk.Entry(self.body, textvariable=v_pct, width=10).grid(row=row, column=1, **pad)
            ttk.Entry(self.body, textvariable=v_min, width=10).grid(row=row, column=2, **pad)
            ttk.Entry(self.body, textvariable=v_max, width=16).grid(row=row, column=3, **pad)
            self.entries[sev] = (v_pct, v_min, v_max)
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
            for sev, (v_pct, v_min, v_max) in self.entries.items():
                max_units = v_max.get().strip()
                self.cfg["severity_map"][sev] = {
                    "pct": float(v_pct.get()),
                    "min_units": int(v_min.get()),
                    "max_units": None if max_units == "" else int(max_units),
                }
            save_config(self.cfg)
            self.on_save(self.cfg)
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Input", "Ensure numeric fields are valid.", parent=self)

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
        ttk.Button(presets, text="Main", command=self._preset_main).pack(side="left", padx=2)
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
        m = self.cfg["severity_map"][severity]
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

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Inventory Qty Comparator — Simple")
        self.root.minsize(820, 460)
        self.theme = ThemeManager(self.root)
        self.results = []
        self.cfg = load_config()
        self.ctx = ContextMenuManager(self.root)
        self.theme.apply_to_toplevel(self.root)
        outer = ttk.Frame(self.root, padding=(10, 8))
        outer.pack(fill="both", expand=True)
        toolbar = ttk.Frame(outer)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Evaluate", style="Accent.TButton", command=self.evaluate_one).pack(side="left")
        ttk.Button(toolbar, text="Settings…", command=self.open_settings).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Choose Columns…", command=self.choose_columns).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Export Excel", command=self.export_excel).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Clear Results", command=self.clear_results).pack(side="left", padx=(6, 0))
        inputs = ttk.Frame(outer)
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
        table_box = ttk.Frame(outer)
        table_box.pack(fill="both", expand=True)
        self.tree_all_columns = list(ALL_COLUMNS)
        self.tree = ttk.Treeview(table_box, columns=self.cfg["visible_columns"], show="headings", height=12)
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
        self.status = ttk.Label(outer, text="Ready.")
        self.status.pack(fill="x", pady=(6, 0))
        self.root.bind("<Return>", lambda e: self.evaluate_one())
        self._toggle_low_stock()
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)

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
        SettingsDialog(self.root, self.theme, self.cfg, on_save)

    def choose_columns(self):
        def on_save(selected_cols: list[str]):
            self.cfg["visible_columns"] = selected_cols
            save_config(self.cfg)
            self._apply_tree_columns(selected_cols)
            self._refresh_tree_rows()
            self.status.configure(text=f"Columns updated ({len(selected_cols)} shown).")
        ColumnChooserDialog(self.root, self.theme, self.cfg.get("visible_columns", DEFAULT_VISIBLE_COLUMNS), on_save)

    def export_excel(self):
        if not self.results:
            messagebox.showinfo("Export", "No results to export yet.")
            return
        default_name = f"inventory_results_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(title="Export to Excel", defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel Workbook", "*.xlsx")])
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

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    App().run()

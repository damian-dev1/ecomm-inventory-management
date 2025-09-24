import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd
from datetime import datetime

APP_BG = "#040111"
PANEL_BG = "#0A0124"
FG = "#ffffff"
ACCENT = "#00d4ff"
TREE_BG = "#040622"
TREE_SEL = "#271b5e"
ROW_STRIPE = "#1a1337"
BTN_BG = "#2a2a2a"
BTN_BG_ACTIVE = "#3a3a3a"
BORDER = "#383a4a"

class ThemeManager:
    def __init__(self, root: tk.Tk):
        self.root = root

    def apply(self):
        s = ttk.Style(self.root)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg=APP_BG)
        # Global option DB for popdown list (critically: dark dropdown)
        self.root.option_add("*TCombobox*Listbox*Background", "#0e0f1c")
        self.root.option_add("*TCombobox*Listbox*Foreground", FG)
        self.root.option_add("*TCombobox*Listbox*selectBackground", TREE_SEL)
        self.root.option_add("*TCombobox*Listbox*selectForeground", FG)
        self.root.option_add("*TCombobox*Listbox*borderWidth", 0)

        s.configure(".", background=APP_BG, foreground=FG)
        s.configure("TFrame", background=APP_BG)
        s.configure("Toolbar.TFrame", background=APP_BG)
        s.configure("Side.TLabelframe", background=PANEL_BG, foreground=FG, bordercolor=BORDER, relief="solid", borderwidth=1)
        s.configure("Side.TLabelframe.Label", background=PANEL_BG, foreground=FG, font=("Segoe UI", 10, "bold"))
        s.configure("TLabel", background=APP_BG, foreground=FG, font=("Segoe UI", 10))
        s.configure("Hint.TLabel", background=PANEL_BG, foreground=ACCENT, font=("Segoe UI", 9))
        s.configure("TButton", background=BTN_BG, foreground=FG, font=("Segoe UI", 10), padding=6, relief="flat")
        s.map("TButton", background=[("active", BTN_BG_ACTIVE)])
        s.configure("Accent.TButton", background=ACCENT, foreground="#001018")
        s.map("Accent.TButton", background=[("active", "#20e1ff")])

        s.configure("Dark.TCheckbutton", background=APP_BG, foreground=FG)
        s.configure("Dark.TSeparator", background=BORDER)

        s.configure("Dark.TCombobox",
                    fieldbackground="#101224",
                    background="#101224",
                    foreground=FG,
                    arrowcolor=FG,
                    selectbackground=TREE_SEL,
                    selectforeground=FG)
        s.map("Dark.TCombobox",
              fieldbackground=[("readonly", "#101224"), ("active", "#14183a")],
              foreground=[("disabled", "#7f7f7f")])

        s.configure("Dark.Treeview",
                    background=TREE_BG,
                    foreground=FG,
                    fieldbackground=TREE_BG,
                    bordercolor=BORDER,
                    rowheight=26)
        s.map("Dark.Treeview",
              background=[("selected", TREE_SEL)],
              foreground=[("selected", FG)])
        s.configure("Dark.Treeview.Heading",
                    background="#202335",
                    foreground=FG,
                    font=("Segoe UI", 10, "bold"),
                    relief="flat")
        s.map("Dark.Treeview.Heading", background=[("active", "#262b48")])

        s.configure("Status.TLabel", background="#090a16", foreground="#b7c2ff")

class StockComparisonApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Stock Comparison Tool")
        self.root.geometry("950x640")
        self.theme = ThemeManager(self.root)
        self.theme.apply()

        self.warehouse_df: pd.DataFrame | None = None
        self.ecommerce_df: pd.DataFrame | None = None

        self.wh_keys = [tk.StringVar() for _ in range(3)]
        self.ec_keys = [tk.StringVar() for _ in range(3)]
        self.show_only_diffs = tk.BooleanVar(value=False)

        self._wh_combos: list[ttk.Combobox] = []
        self._ec_combos: list[ttk.Combobox] = []

        self._build_ui()
        self._bind_keys()

    def _bind_keys(self):
        self.root.bind("<Control-o>", lambda e: self._menu_load_csv("warehouse"))
        self.root.bind("<Control-e>", lambda e: self._menu_load_csv("ecommerce"))
        self.root.bind("<Control-r>", lambda e: self.compare_data())

    def _build_ui(self):
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        toolbar = ttk.Frame(self.root, style="Toolbar.TFrame", padding=(10, 10, 10, 6))
        toolbar.grid(row=0, column=0, sticky="ew")
        for i in range(6):
            toolbar.grid_columnconfigure(i, weight=0)
        toolbar.grid_columnconfigure(6, weight=1)

        ttk.Button(toolbar, text="Load Warehouse (Ctrl+O)", command=lambda: self._menu_load_csv("warehouse")).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(toolbar, text="Load Ecommerce (Ctrl+E)", command=lambda: self._menu_load_csv("ecommerce")).grid(row=0, column=1, padx=6)
        ttk.Separator(toolbar, orient="vertical", style="Dark.TSeparator").grid(row=0, column=2, sticky="ns", padx=8)
        ttk.Checkbutton(toolbar, text="Show only differences", variable=self.show_only_diffs, style="Dark.TCheckbutton", command=self.compare_data).grid(row=0, column=3, padx=6)
        ttk.Button(toolbar, text="Compare (Ctrl+R)", style="Accent.TButton", command=self.compare_data).grid(row=0, column=4, padx=6)
        ttk.Button(toolbar, text="Export Visible CSV", command=self._export_visible).grid(row=0, column=5, padx=(6, 0))
        ttk.Label(toolbar, text=" ").grid(row=0, column=6, sticky="ew")  # spacer

        main = ttk.Frame(self.root, padding=(10, 6, 10, 10))
        main.grid(row=1, column=0, sticky="nsew")
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=0)
        main.grid_columnconfigure(1, weight=1)

        side = ttk.Labelframe(main, text="Mappings", padding=10, style="Side.TLabelframe")
        side.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        side.grid_columnconfigure(0, weight=1)

        wh_frame = ttk.Labelframe(side, text="Warehouse", padding=10, style="Side.TLabelframe")
        wh_frame.grid(row=0, column=0, sticky="ew")
        self._wh_combos = self._add_mapping_rows(wh_frame, self.wh_keys)

        ec_frame = ttk.Labelframe(side, text="Ecommerce", padding=10, style="Side.TLabelframe")
        ec_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self._ec_combos = self._add_mapping_rows(ec_frame, self.ec_keys)

        results = ttk.Frame(main)
        results.grid(row=0, column=1, sticky="nsew")
        results.grid_rowconfigure(0, weight=1)
        results.grid_columnconfigure(0, weight=1)

        cols = ("Key", "Warehouse Qty", "Ecommerce Qty", "Difference")
        self.tree = ttk.Treeview(results, columns=cols, show="headings", style="Dark.Treeview")
        for c in cols:
            self.tree.heading(c, text=c, command=lambda col=c: self._sort_tree_by(col, False))
            width = 300 if c == "Key" else 150
            self.tree.column(c, width=width, anchor="w" if c == "Key" else "e", stretch=True)

        yscroll = ttk.Scrollbar(results, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(results, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure("even", background=TREE_BG)
        self.tree.tag_configure("odd", background=ROW_STRIPE)
        self.tree.tag_configure("pos", foreground="#86ff9c")
        self.tree.tag_configure("neg", foreground="#ff9c9c")
        self.tree.tag_configure("zero", foreground="#c8c8c8")

        self._build_tree_menu()

        status = ttk.Frame(self.root, padding=(10, 6))
        status.grid(row=2, column=0, sticky="ew")
        status.grid_columnconfigure(0, weight=1)
        self.status_label = ttk.Label(status, text="Ready", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w")

    def _build_tree_menu(self):
        self.tree_menu = tk.Menu(self.root, tearoff=0, bg="#0e0f1c", fg=FG, activebackground=TREE_SEL, activeforeground=FG, bd=0)
        self.tree_menu.add_command(label="Copy Cell", command=self._copy_cell)
        self.tree_menu.add_command(label="Copy Row", command=self._copy_row)
        self.tree.bind("<Button-3>", self._open_tree_menu)

    def _open_tree_menu(self, event):
        rowid = self.tree.identify_row(event.y)
        if rowid:
            self.tree.selection_set(rowid)
            self.tree_menu.tk_popup(event.x_root, event.y_root)

    def _copy_cell(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        col = self.tree.identify_column(self.tree.winfo_pointerx() - self.tree.winfo_rootx())
        try:
            cidx = int(col.replace("#", "")) - 1
        except Exception:
            cidx = 0
        val = self.tree.item(iid, "values")[cidx]
        self.root.clipboard_clear()
        self.root.clipboard_append(str(val))

    def _copy_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        self.root.clipboard_clear()
        self.root.clipboard_append("\t".join(map(str, vals)))

    def _add_mapping_rows(self, parent: ttk.Labelframe, key_vars: list[tk.StringVar]) -> list[ttk.Combobox]:
        combos: list[ttk.Combobox] = []
        for i in range(3):
            row = ttk.Frame(parent)
            row.grid(row=i, column=0, sticky="ew", pady=(0, 8))
            ttk.Label(row, text=f"Key {i+1}:").pack(side="left")
            cb = ttk.Combobox(row, state="readonly", width=28, textvariable=key_vars[i], style="Dark.TCombobox")
            cb.pack(side="left", padx=(8, 0), fill="x", expand=True)
            cb.set("")
            combos.append(cb)
        ttk.Label(parent, text="Key 1 = ID, Key 3 = Quantity", style="Hint.TLabel").grid(row=3, column=0, sticky="w", pady=(4, 0))
        return combos

    @staticmethod
    def _sniff_delimiter(path: Path) -> str:
        with path.open("r", encoding="utf-8", errors="replace") as f:
            sample = f.read(4096)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            return dialect.delimiter
        except csv.Error:
            return ","

    def _menu_load_csv(self, source: str):
        self.load_csv(source)

    def load_csv(self, source: str):
        fpath = filedialog.askopenfilename(
            title=f"Select {source.capitalize()} CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not fpath:
            return
        p = Path(fpath)
        delim = self._sniff_delimiter(p)
        try:
            df = pd.read_csv(p, dtype=str, sep=delim, low_memory=False)
        except Exception as e:
            messagebox.showerror("Load CSV", f"Failed to load file:\n{e}")
            return
        if df.empty:
            messagebox.showwarning("Load CSV", "The selected CSV is empty.")
            return

        cols = list(df.columns)
        if not cols:
            messagebox.showwarning("Load CSV", "No columns detected in CSV.")
            return

        if source == "warehouse":
            self.warehouse_df = df
            combos = self._wh_combos
            targets = self.wh_keys
        else:
            self.ecommerce_df = df
            combos = self._ec_combos
            targets = self.ec_keys

        for cb in combos:
            cb["values"] = cols

        default_ids = [c for c in cols if c.lower() in {"sku", "id", "product_id"}]
        default_qty = [c for c in cols if ("qty" in c.lower()) or ("quantity" in c.lower()) or (c.lower() in {"stock", "soh"})]

        k1 = (default_ids[0] if default_ids else cols[0])
        k3 = (default_qty[0] if default_qty else (cols[1] if len(cols) > 1 else cols[0]))
        targets[0].set(k1)
        targets[1].set("")  # optional
        targets[2].set(k3)

        messagebox.showinfo("Loaded", f"{source.capitalize()} CSV loaded with {len(df):,} rows and {len(cols)} columns.")
        self._set_status(f"{source.capitalize()} file: {p.name} | rows={len(df):,}")

    def compare_data(self):
        self.tree.delete(*self.tree.get_children())
        if self.warehouse_df is None or self.ecommerce_df is None:
            messagebox.showwarning("Compare", "Load both Warehouse and Ecommerce CSVs first.")
            return

        wh_id = self.wh_keys[0].get().strip()
        wh_qty = self.wh_keys[2].get().strip()
        ec_id = self.ec_keys[0].get().strip()
        ec_qty = self.ec_keys[2].get().strip()
        if not all([wh_id, wh_qty, ec_id, ec_qty]):
            messagebox.showwarning("Compare", "Select ID (Key 1) and Quantity (Key 3) for both sides.")
            return

        try:
            wh = self.warehouse_df[[wh_id, wh_qty]].dropna(subset=[wh_id]).copy()
            ec = self.ecommerce_df[[ec_id, ec_qty]].dropna(subset=[ec_id]).copy()

            wh[wh_id] = wh[wh_id].astype(str).str.strip()
            ec[ec_id] = ec[ec_id].astype(str).str.strip()

            wh = wh.drop_duplicates(subset=[wh_id], keep="last").set_index(wh_id)
            ec = ec.drop_duplicates(subset=[ec_id], keep="last").set_index(ec_id)

            wh[wh_qty] = pd.to_numeric(wh[wh_qty], errors="coerce").fillna(0)
            ec[ec_qty] = pd.to_numeric(ec[ec_qty], errors="coerce").fillna(0)

            keys = wh.index.intersection(ec.index)
            inserted = 0
            only_diffs = self.show_only_diffs.get()

            for i, key in enumerate(keys):
                wq = float(wh.loc[key, wh_qty])
                eq = float(ec.loc[key, ec_qty])
                diff = wq - eq
                if only_diffs and diff == 0:
                    continue
                stripe = "odd" if i % 2 else "even"
                sign = "zero" if diff == 0 else ("pos" if diff > 0 else "neg")
                self.tree.insert("", "end",
                                 values=(key, f"{wq:.0f}", f"{eq:.0f}", f"{diff:.0f}"),
                                 tags=(stripe, sign))
                inserted += 1

            if inserted == 0:
                messagebox.showinfo("Compare", "No rows to display (check mappings or toggle).")

            self._set_status(f"Compared: {len(keys):,} matching IDs | displayed: {inserted:,} | {datetime.now().strftime('%H:%M:%S')}")
        except Exception as e:
            messagebox.showerror("Compare", f"Error comparing data:\n{e}")

    def _sort_tree_by(self, col_name: str, descending: bool):
        columns = self.tree["columns"]
        idx = columns.index(col_name)
        rows = [(self.tree.set(k, col_name), k) for k in self.tree.get_children("")]
        def _to_num(v):
            try:
                return float(str(v).replace(",", ""))
            except Exception:
                return None
        if col_name in ("Warehouse Qty", "Ecommerce Qty", "Difference"):
            decorated = [((_to_num(v) if _to_num(v) is not None else float("inf")), k) for v, k in rows]
        else:
            decorated = [(str(v).lower(), k) for v, k in rows]
        decorated.sort(reverse=descending)
        for i, (_, k) in enumerate(decorated):
            self.tree.move(k, "", i)
        self.tree.heading(col_name, command=lambda: self._sort_tree_by(col_name, not descending))

    def _export_visible(self):
        items = self.tree.get_children("")
        if not items:
            messagebox.showinfo("Export", "No rows to export.")
            return
        save = filedialog.asksaveasfilename(
            title="Export visible rows",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not save:
            return
        cols = self.tree["columns"]
        try:
            with open(save, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(cols)
                for iid in items:
                    writer.writerow(self.tree.item(iid, "values"))
            messagebox.showinfo("Export", f"Exported {len(items):,} rows to:\n{save}")
        except Exception as e:
            messagebox.showerror("Export", f"Failed to export:\n{e}")

    def _set_status(self, msg: str):
        self.status_label.config(text=msg)

def main():
    root = tk.Tk()
    app = StockComparisonApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

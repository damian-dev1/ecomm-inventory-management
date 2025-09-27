import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog, StringVar
import gc

class SKUComparerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Vendor SKU Comparator")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- Data Storage ---
        self.df_a = None
        self.df_b = None
        self.comparison_results = pd.DataFrame()
        self.current_view_df = pd.DataFrame()

        # --- Main Layout ---
        top_frame = ttk.Frame(root)
        top_frame.pack(side="top", fill="both", expand=True)
        
        # --- Status Bar ---
        self.status_var = StringVar(value="Ready")
        status_bar = ttk.Frame(root, padding=(5, 2), bootstyle=SECONDARY)
        status_bar.pack(side="bottom", fill="x")
        status_label = ttk.Label(status_bar, textvariable=self.status_var)
        status_label.pack(side="left")

        left_frame = ttk.Frame(top_frame, padding=10)
        left_frame.pack(side="left", fill="y")

        right_frame = ttk.Frame(top_frame, padding=(0, 10, 10, 10))
        right_frame.pack(side="left", fill="both", expand=True)

        # --- Tabs ---
        self.tabs = ttk.Notebook(right_frame)
        self.tabs.pack(fill="both", expand=True)
        self.live_tab = ttk.Frame(self.tabs)
        self.preview_tab = ttk.Frame(self.tabs)
        self.summary_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.live_tab, text="🔴 Live")
        self.tabs.add(self.preview_tab, text="🔍 Preview")
        self.tabs.add(self.summary_tab, text="📈 Summary")
        
        # --- UI Elements ---
        self.create_stats_sidebar(left_frame)
        self.create_live_tab(self.live_tab)
        self.create_preview_tab(self.preview_tab)

        self.log("💡 Welcome! Load two datasets to enable the 'Compare' button.")
        self.update_status("Ready")

    def create_stats_sidebar(self, parent):
        stats_frame = ttk.LabelFrame(parent, text="📊 Comparison Stats", padding=10)
        stats_frame.pack(fill="x")
        self.stats_labels = {}
        stats_to_create = [
            ("Source A Rows", "N/A"), ("Source B Rows", "N/A"),
            ("Source A Unique SKUs", "N/A"), ("Source B Unique SKUs", "N/A"),
            ("---", "---"),
            ("✅ Matched", "0"), ("❌ Mismatched", "0"),
            ("🅰️ Only in A", "0"), ("🅱️ Only in B", "0"),
            ("---", "---"),
            ("📈 Match Rate", "0%"), ("📉 Qty Variance", "0")
        ]
        for i, (text, val) in enumerate(stats_to_create):
            if text == "---":
                ttk.Separator(stats_frame, orient=HORIZONTAL).grid(row=i, column=0, columnspan=2, pady=5, sticky="ew")
                continue
            ttk.Label(stats_frame, text=f"{text}:").grid(row=i, column=0, sticky="w", pady=2)
            self.stats_labels[text] = ttk.Label(stats_frame, text=val, font=("Helvetica", 10, "bold"))
            self.stats_labels[text].grid(row=i, column=1, sticky="e", padx=(10, 0))

    def create_live_tab(self, parent):
        paned_window = ttk.PanedWindow(parent, orient=VERTICAL)
        paned_window.pack(fill=BOTH, expand=True, pady=5)
        
        # --- Top Pane: Controls ---
        control_pane = ttk.Frame(paned_window, padding=5)
        # FIX: Set weight to 0 so this pane only takes the space it needs initially.
        paned_window.add(control_pane, weight=0)
        
        dataset_frame = ttk.Frame(control_pane)
        # FIX: Set expand=False to prevent the frame from growing vertically.
        dataset_frame.pack(fill="x", pady=5, expand=False)
        
        a_frame = ttk.LabelFrame(dataset_frame, text="📂 Dataset A")
        # FIX: Set fill='x' to prevent vertical expansion of the inner labelframe.
        a_frame.pack(side="left", fill="x", expand=True, padx=5)
        self.headers_a = self.create_dataset_controls(a_frame, "A")
        
        b_frame = ttk.LabelFrame(dataset_frame, text="📁 Dataset B")
        # FIX: Set fill='x' to prevent vertical expansion of the inner labelframe.
        b_frame.pack(side="left", fill="x", expand=True, padx=5)
        self.headers_b = self.create_dataset_controls(b_frame, "B")

        action_frame = ttk.Frame(control_pane)
        action_frame.pack(pady=10)
        self.compare_btn = ttk.Button(action_frame, text="🔍 Compare Datasets", command=self.compare_datasets, bootstyle=SUCCESS, state=DISABLED)
        self.compare_btn.pack(side="left", padx=5)
        ToolTip(self.compare_btn, text="Run comparison (enabled after loading both datasets)")
        export_btn = ttk.Button(action_frame, text="📤 Export Current View", command=self.export_results, bootstyle=INFO)
        export_btn.pack(side="left", padx=5)
        ToolTip(export_btn, text="Save the currently displayed rows to a CSV file")

        # --- Bottom Pane: Logs ---
        logs_pane = ttk.LabelFrame(paned_window, text="📜 Logs", padding=5)
        # FIX: Set weight to 1 so the logs take all remaining vertical space.
        paned_window.add(logs_pane, weight=1)
        self.log_text = ttk.Text(logs_pane, height=10, wrap="word", font=("Courier New", 9))
        self.log_text.pack(fill="both", expand=True)

    def create_dataset_controls(self, frame, label):
        load_btn = ttk.Button(frame, text=f"📂 Load Dataset {label}", command=lambda: self.load_dataset(label), bootstyle=PRIMARY)
        load_btn.pack(pady=10, padx=10, fill="x")
        ToolTip(load_btn, text=f"Select a CSV file for Dataset {label}")
        headers = {}
        for col in ["Vendor ID", "SKU", "Qty"]:
            ttk.Label(frame, text=f"{col} Column:").pack(padx=10, anchor="w")
            headers[col] = ttk.Combobox(frame, state="readonly", width=18)
            headers[col].pack(pady=(0, 5), padx=10, fill="x")
            ToolTip(headers[col], text=f"Select column for {col}")
        return headers

    def create_preview_tab(self, parent):
        controls_frame = ttk.Frame(parent)
        controls_frame.pack(fill="x", padx=5, pady=5)
        self.view_var = StringVar(value="Mismatched")
        view_options = ["Mismatched", "Only in A", "Only in B", "All Results"]
        for option in view_options:
            rb = ttk.Radiobutton(controls_frame, text=option, variable=self.view_var, value=option, command=self.update_view)
            rb.pack(side="left", padx=5)
        ttk.Separator(controls_frame, orient=VERTICAL).pack(side="left", fill='y', padx=10)
        ttk.Label(controls_frame, text="Filter by Vendor ID:").pack(side="left", padx=5)
        self.vendor_filter = ttk.Entry(controls_frame, width=15)
        self.vendor_filter.pack(side="left", padx=5)
        ttk.Label(controls_frame, text="Filter by SKU:").pack(side="left", padx=5)
        self.sku_filter = ttk.Entry(controls_frame, width=15)
        self.sku_filter.pack(side="left", padx=5)
        apply_filter_btn = ttk.Button(controls_frame, text="Apply Filter", command=self.update_view, bootstyle=INFO)
        apply_filter_btn.pack(side="left", padx=5)
        
        self.preview_limit_label = ttk.Label(parent, text="", bootstyle=WARNING)
        self.preview_limit_label.pack(fill='x', padx=5, pady=(5,0))

        cols = ("Status", "Vendor ID", "SKU", "Qty A", "Qty B")
        self.tree = ttk.Treeview(parent, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c, False))
            self.tree.column(col, width=120, anchor="center")
        self.tree.tag_configure("mismatch", background="#3e2c2c", foreground="#ff8a8a")
        self.tree.tag_configure("only_a", background="#4a4a2a", foreground="#ffffa8")
        self.tree.tag_configure("only_b", background="#2a414a", foreground="#a8deff")
        self.tree.pack(fill="both", expand=True, padx=5, pady=(0,5))

    def update_status(self, message, clear_after_ms=0):
        self.status_var.set(message)
        if clear_after_ms > 0:
            self.root.after(clear_after_ms, lambda: self.status_var.set("Ready") if self.status_var.get() == message else None)

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def load_dataset(self, label):
        file_path = filedialog.askopenfilename(title=f"Select Dataset {label} CSV", filetypes=[("CSV Files", "*.csv")])
        if not file_path: return
        self.update_status(f"Loading {label}: {file_path.split('/')[-1]}...")
        try:
            df = pd.read_csv(file_path)
            if label == "A":
                self.df_a = df
                self.update_headers_combobox(self.headers_a, df.columns)
            else:
                self.df_b = df
                self.update_headers_combobox(self.headers_b, df.columns)
            self.log(f"✅ Dataset {label} loaded successfully with {len(df)} rows.")
            self.update_status(f"Dataset {label} loaded successfully.", 5000)
        except Exception as e:
            self.log(f"❌ Error loading Dataset {label}: {e}")
            self.update_status(f"Error loading Dataset {label}.", 5000)
            Messagebox.show_error(f"Could not read the file.\n\nError: {e}", f"File Load Error: Dataset {label}")
        
        if self.df_a is not None and self.df_b is not None:
            self.compare_btn.config(state=NORMAL)
            self.log("✅ Both datasets loaded. Ready to compare.")
        else:
            self.compare_btn.config(state=DISABLED)

    def update_headers_combobox(self, header_widgets, columns):
        col_list = list(columns)
        for widget in header_widgets.values():
            widget["values"] = col_list
            if col_list: widget.set(col_list[0])

    def validate_and_prepare_df(self, df, column_map, df_name):
        df_copy = df.copy()
        for key_col, (app_col, dtype) in column_map.items():
            if dtype == 'int':
                original_series = df_copy[app_col]
                numeric_series = pd.to_numeric(original_series, errors='coerce')
                failed_rows = numeric_series.isna() & original_series.notna()
                if failed_rows.sum() > 0:
                    raise ValueError(f"Column '{app_col}' in Dataset {df_name} contains non-integer values.")
                df_copy[key_col] = numeric_series.fillna(0).astype('int64')
            else: # 'str'
                df_copy[key_col] = df_copy[app_col].astype(str)
        return df_copy[[c for c in column_map.keys()]]

    def compare_datasets(self):
        self.update_status("Starting comparison...")
        self.log("🚀 Starting comparison...")
        try:
            cols_a = {'vendor_id': (self.headers_a["Vendor ID"].get(), 'int'), 'sku': (self.headers_a["SKU"].get(), 'str'), 'qty_A': (self.headers_a["Qty"].get(), 'int')}
            cols_b = {'vendor_id': (self.headers_b["Vendor ID"].get(), 'int'), 'sku': (self.headers_b["SKU"].get(), 'str'), 'qty_B': (self.headers_b["Qty"].get(), 'int')}
            df1 = self.validate_and_prepare_df(self.df_a, cols_a, "A")
            df2 = self.validate_and_prepare_df(self.df_b, cols_b, "B")
            self.log("✅ Data validation successful.")
            self.update_status("Merging datasets...")
            merged = pd.merge(df1, df2, on=["vendor_id", "sku"], how="outer", indicator=True)
            merged.fillna({'qty_A': 0, 'qty_B': 0}, inplace=True)
            merged['qty_A'] = merged['qty_A'].astype(int)
            merged['qty_B'] = merged['qty_B'].astype(int)
            def get_status(row):
                if row['_merge'] == 'left_only': return 'Only in A'
                if row['_merge'] == 'right_only': return 'Only in B'
                return 'Matched' if row['qty_A'] == row['qty_B'] else 'Mismatched'
            merged['Status'] = merged.apply(get_status, axis=1)
            self.comparison_results = merged.drop(columns=['_merge'])
            self.log("✅ Comparison complete.")
            self.update_status("Comparison complete.", 5000)
            self.update_all_stats()
            self.update_view()
            self.update_summary_chart()
            self.tabs.select(self.preview_tab)
        except ValueError as ve:
            self.log(f"❌ VALIDATION ERROR: {ve}")
            self.update_status(f"Validation Error! See logs.", 10000)
            Messagebox.show_error(str(ve), "Data Validation Error")
        except Exception as e:
            self.log(f"❌ An unexpected error occurred during comparison: {e}")
            self.update_status(f"Comparison failed! See logs.", 10000)
            Messagebox.show_error(f"An unexpected error occurred.\n\nError: {e}", "Comparison Failed")

    def update_all_stats(self):
        if self.comparison_results.empty: return
        stats = {
            "Source A Rows": len(self.df_a), "Source B Rows": len(self.df_b),
            "Source A Unique SKUs": self.df_a[self.headers_a["SKU"].get()].nunique(),
            "Source B Unique SKUs": self.df_b[self.headers_b["SKU"].get()].nunique(),
            "✅ Matched": (self.comparison_results['Status'] == 'Matched').sum(),
            "❌ Mismatched": (self.comparison_results['Status'] == 'Mismatched').sum(),
            "🅰️ Only in A": (self.comparison_results['Status'] == 'Only in A').sum(),
            "🅱️ Only in B": (self.comparison_results['Status'] == 'Only in B').sum()
        }
        total_common = stats["✅ Matched"] + stats["❌ Mismatched"]
        stats["📈 Match Rate"] = f"{(stats['✅ Matched'] / total_common * 100) if total_common > 0 else 0:.2f}%"
        mismatched_df = self.comparison_results[self.comparison_results['Status'] == 'Mismatched']
        stats["📉 Qty Variance"] = int((mismatched_df['qty_B'] - mismatched_df['qty_A']).sum())
        for key, val in stats.items():
            if key in self.stats_labels: self.stats_labels[key].config(text=str(val))

    # FIX: Added 'self' as the first parameter to the method definition.
    def update_view(self):
        if self.comparison_results.empty: return
        view_filter = self.view_var.get()
        df = self.comparison_results.copy()
        if view_filter != "All Results": df = df[df["Status"] == view_filter]
        vendor = self.vendor_filter.get().strip().lower()
        sku = self.sku_filter.get().strip().lower()
        if vendor: df = df[df["vendor_id"].astype(str).str.lower().str.contains(vendor, na=False)]
        if sku: df = df[df["sku"].astype(str).str.lower().str.contains(sku, na=False)]
        self.current_view_df = df
        self.update_treeview(df)

    def update_treeview(self, df):
        self.tree.delete(*self.tree.get_children())
        
        total_rows = len(df)
        df_to_display = df

        if total_rows > 100:
            self.preview_limit_label.config(text=f"⚠️ Displaying first 100 of {total_rows} rows. Export for full results.")
            df_to_display = df.head(100)
        else:
            self.preview_limit_label.config(text="")
        
        tag_map = {"Mismatched": "mismatch", "Only in A": "only_a", "Only in B": "only_b"}
        for _, row in df_to_display.iterrows():
            tags = (tag_map.get(row["Status"], ""),)
            values = (row["Status"], row["vendor_id"], row["sku"], int(row["qty_A"]), int(row["qty_B"]))
            self.tree.insert("", "end", values=values, tags=tags)
    def sort_treeview(self, col: str, reverse: bool):
        # Collect current values for the target column
        rows = []
        for iid in self.tree.get_children(""):
            v = self.tree.set(iid, col)
            rows.append((v, iid))

        def _num_or_none(x: str):
            try:
                # Handle empty/None gracefully
                return float(x.replace(",", "")) if isinstance(x, str) else float(x)
            except Exception:
                return None

        # Decide numeric vs lexicographic based on presence of parsable numbers
        numeric_values = [n for n, _ in (( _num_or_none(v), iid) for v, iid in rows) if n is not None]
        if len(numeric_values) == len(rows):  # all values numeric
            rows.sort(key=lambda t: _num_or_none(t[0]), reverse=reverse)
        else:
            # Case-insensitive string sort, None/"" last
            rows.sort(key=lambda t: (t[0] in (None, ""), str(t[0]).lower()), reverse=reverse)

        # Reorder items
        for idx, (_, iid) in enumerate(rows):
            self.tree.move(iid, "", idx)

        # Toggle sort order next click
        self.tree.heading(col, command=lambda c=col: self.sort_treeview(c, not reverse))
    def update_summary_chart(self):
        # Clear previous widgets
        for w in self.summary_tab.winfo_children():
            w.destroy()

        if self.comparison_results is None or self.comparison_results.empty:
            return

        # Stable order for labels
        order = ["Matched", "Mismatched", "Only in A", "Only in B"]
        vc = self.comparison_results["Status"].value_counts().reindex(order, fill_value=0)
        labels = list(vc.index)
        sizes = vc.to_numpy()

        # Colors
        color_map = {
            "Matched": "#28a745",
            "Mismatched": "#dc3545",
            "Only in A": "#ffc107",
            "Only in B": "#17a2b8",
        }
        pie_colors = [color_map.get(lbl, "#6c757d") for lbl in labels]

        # Theme-aware background/foreground
        try:
            style = getattr(self.root, "style", ttk.Style())
        except Exception:
            style = ttk.Style()

        bg = getattr(getattr(style, "colors", None), "bg", None)
        fg = getattr(getattr(style, "colors", None), "fg", None)

        if not bg:
            bg = style.lookup("TFrame", "background") or "#ffffff"
        if not fg:
            # Simple luminance check to choose a contrasting text color
            def _lum(hexcolor: str):
                hexcolor = hexcolor.lstrip("#")
                r, g, b = int(hexcolor[0:2], 16), int(hexcolor[2:4], 16), int(hexcolor[4:6], 16)
                return 0.2126*r + 0.7152*g + 0.0722*b
            fg = "#ffffff" if _lum(bg) < 128 else "#000000"

        fig, ax = plt.subplots(figsize=(5, 4), dpi=100)
        fig.patch.set_facecolor(bg)
        ax.set_facecolor(bg)

        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=labels,
            autopct="%1.1f%%",
            startangle=90,
            colors=pie_colors,
            wedgeprops={"edgecolor": fg, "linewidth": 0.5},
            textprops={"color": fg},
        )
        ax.axis("equal")
        ax.set_title("Comparison Breakdown", color=fg)

        canvas = FigureCanvasTkAgg(fig, master=self.summary_tab)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

    def export_results(self):
        if self.current_view_df.empty:
            self.log("⚠️ No data in the current view to export.")
            self.update_status("No data to export.", 5000)
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")], title="Save Current View")
        if file_path:
            self.update_status(f"Exporting full results for view to {file_path.split('/')[-1]}...")
            try:
                self.current_view_df.to_csv(file_path, index=False)
                self.log(f"📁 Full view ({len(self.current_view_df)} rows) exported successfully to: {file_path}")
                self.update_status("Export successful.", 5000)
            except Exception as e:
                self.log(f"❌ Error exporting data: {e}")
                self.update_status("Export failed.", 5000)

    def on_close(self):
        self.log("🛑 Application is closing...")
        plt.close('all')
        self.df_a = self.df_b = self.comparison_results = self.current_view_df = None
        gc.collect()
        self.root.destroy()

if __name__ == "__main__":
    root = ttk.Window(themename="darkly", size=(1200, 800), minsize=(1000, 650))
    app = SKUComparerApp(root)
    root.mainloop()

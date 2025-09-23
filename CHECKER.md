# Inventory Qty Comparator

## Key Features

* **Primary Stat:** **Stock Qty Comparison** (Match/Mismatch) with reason.
* **Auto Severity** only (no manual severity pickers).
* **Strict vs Tolerant** qty comparison modes (configurable in Settings).
* **Low Stock logic** (optional): status = Low Stock for small quantities.
* **Compact dark UI** with right-click **Quick Paste** in text/entry fields.
* **Column Chooser** for a slim, operator-friendly results grid.
* **One-key run**: press **Enter** to evaluate.
* **Excel export (.xlsx)** including the full data model.
* **Persistent settings** via `tolerance_config.json`.

### Minimal Workflow

1. Type **SKU**, **EC Qty**, **WH Qty**.
2. Optionally tick **Low Stock** and set **Max**.
3. Press **Enter** or click **Evaluate**.
4. Review **Stock Qty Comparison** + **Reason**.
5. Export via **Export Excel** if needed.

---

## UI Tour

* **Toolbar**

  * **Evaluate**: run the check.
  * **Settings…**: choose compare mode & tolerance mechanics.
  * **Choose Columns…**: show only what matters.
  * **Export Excel**: save a full report.
  * **Clear Results**: wipe session rows.

* **Inputs**

  * **SKU** (free text)
  * **EC Qty** (integer)
  * **WH Qty** (integer)
  * **Low Stock** + **Max** (enables “Low Stock” status when qty ≤ Max)

* **Results Grid**

  * Color tags for quick scanning:

    * **Match** rows show in success color
    * **Mismatch** rows show in error color
  * Choose which columns are visible (compact by default)

* **Status Bar**

  * Shows last evaluation summary (Δ and tolerance units).

* **Context Menus**

  * Right-click any Entry/Text for **Cut/Copy/Paste/Select All**.

---

## Data Model (Row Fields)

| Column                   | Description                                     |              |            |           |             |
| ------------------------ | ----------------------------------------------- | ------------ | ---------- | --------- | ----------- |
| Timestamp                | Evaluation time (local)                         |              |            |           |             |
| SKU                      | Identifier provided in input                    |              |            |           |             |
| Ecommerce Qty            | Integer input                                   |              |            |           |             |
| Warehouse Qty            | Integer input                                   |              |            |           |             |
| Use Low Stock Logic      | True/False                                      |              |            |           |             |
| Low Stock Max Qty        | Threshold for “Low Stock”                       |              |            |           |             |
| Applied Severity         | Auto (derived)                                  |              |            |           |             |
| Tolerance (%)            | Severity’s fractional rate (e.g., 0.10 for 10%) |              |            |           |             |
| Tolerance Base           | Base used: \`max                                | min          | avg        | ecommerce | warehouse\` |
| Tolerance Rounding       | \`floor                                         | ceil         | round\`    |           |             |
| Tolerance Units          | Final integer tolerance (after min/max caps)    |              |            |           |             |
| Delta                    | `abs(EC - WH)`                                  |              |            |           |             |
| Ecommerce Status         | \`Out of Stock                                  | Low Stock    | In Stock\` |           |             |
| Warehouse Status         | \`Out of Stock                                  | Low Stock    | In Stock\` |           |             |
| Availability             | Natural text explanation of combined statuses   |              |            |           |             |
| **Stock Qty Comparison** | \*\*Match                                       | Mismatch\*\* |            |           |             |
| Reason                   | Human-readable explanation                      |              |            |           |             |

> Default **visible columns** are a compact subset focused on decision-making; customize via **Choose Columns…**.

---

## Decision Logic (Short, Sharp, Auditable)

### 1) Status Computation

```
if qty == 0 → "Out of Stock"
elif Use Low Stock AND 1 ≤ qty ≤ LowStockMax → "Low Stock"
else → "In Stock"
```

### 2) Auto Severity (no manual input)

```
if exactly one of the statuses is "Out of Stock" → "Critical"
elif statuses differ → "High"
else (statuses equal) → use config "auto_equal_status_severity" (default "Medium")
```

### 3) Tolerance Units

* Start with `pct * base`, where:

  * `base` = one of `max(ec,wh) | min(ec,wh) | avg(ec,wh) | ec | wh`
  * `pct` from severity map
* Apply rounding: `floor | ceil | round`
* Apply min/max caps: `min_units`, `max_units` (None = no upper bound)

### 4) Qty Compare Mode

* **strict**:

  * **Match** only if `EC Qty == WH Qty`
  * Else **Mismatch** (fast, unforgiving)
* **tolerant**:

  * If statuses differ **and tolerance is 0** → **Mismatch**
  * Else compare `Delta = abs(EC - WH)` to `Tolerance Units`:

    * `Delta ≤ Tol` → **Match**
    * `Delta > Tol` → **Mismatch**

### Example

* EC=100, WH=102, Mode=tolerant, Severity=Medium (5%), Base=max=102

  * Raw tol = 0.05 × 102 = 5.1 → rounding=floor → 5 units
  * Δ = 2 ≤ 5 → **Match** with reason “Delta ≤ tolerance (2 ≤ 5).”

> In **strict** mode the same example would be **Mismatch** because quantities differ.

---

## Settings

* **Qty Compare Mode**: `strict` or `tolerant`
* **Base Method**: `max | avg | min | ecommerce | warehouse`
* **Rounding**: `floor | ceil | round`
* **Auto (equal statuses)**: Which severity to use when statuses are equal (default **Medium**).
* **Severity Map** (read-only in this simplified app’s README; still persisted): per-severity `pct`, `min_units`, `max_units`.

Settings persist in `tolerance_config.json` next to the app.

---

## Export

* Click **Export Excel** to write an `.xlsx`:

  * One sheet named **Results**
  * Header row = full data model (all columns)
  * One row per evaluation
* Works well for **audit trails**, **attachments**, and **stakeholder proof**.

---

## Usage Tips

* **Enter**: runs Evaluate.
* Right-click in any input to **Paste** quickly.
* Use **Choose Columns…** to declutter the view (e.g., when operators only need the decision fields).
* Prefer **strict** mode for hard reconciliation; switch to **tolerant** only when you have a defensible tolerance policy.

---

## Outcomes You Can Expect

* **Consistent** pass/fail decisions across operators.
* **Explainable** mismatch reasons (clean audit trails).
* **Fewer false alarms** when using **tolerant** mode with controlled caps.
* **Faster reviews** via compact, dark, keyboard-friendly UI.
* **Shareable evidence** via Excel export.

---

## Configuration File

`./tolerance_config.json` (auto-created/merged)

Example keys of interest:

```json
{
  "qty_compare_mode": "strict",
  "base_method": "max",
  "rounding": "floor",
  "auto_equal_status_severity": "Medium",
  "severity_map": {
    "Critical": {"pct": 0.0, "min_units": 0, "max_units": 0},
    "High":     {"pct": 0.0, "min_units": 0, "max_units": 0},
    "Medium":   {"pct": 0.05, "min_units": 1, "max_units": null},
    "Low":      {"pct": 0.10, "min_units": 1, "max_units": null},
    "Very Low": {"pct": 0.20, "min_units": 1, "max_units": null}
  },
  "visible_columns": [
    "Timestamp","SKU","Ecommerce Qty","Warehouse Qty",
    "Delta","Applied Severity","Stock Qty Comparison","Reason"
  ]
}
```

---

## Troubleshooting

* **Dropdown not dark?**
  Tk on some platforms/themes can be stubborn; the app forces a best-effort dark popdown via option DB. It’s still readable on Windows/macOS/Linux.
* **No Excel export option in dialog?**
  Ensure file type filter is set to **Excel Workbook**; the app writes `.xlsx` only.
* **Validation errors**
  EC/WH quantities must be integers. The app will prompt if they aren’t.

---

## Roadmap (kept lean)

* Batch input panel (CSV) for bulk per-row evaluation (optional).
* Hotkeys for export/columns/settings.
* Pluggable **severity profiles** (kept off by default to preserve simplicity).

---

## License

Proprietary / Internal use.

---

## Credits

Built with **Tkinter** and **openpyxl**. Dark theme inspired by Tokyo-Night palettes.

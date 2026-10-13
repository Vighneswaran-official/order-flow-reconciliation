# Order Flow & Settlement Reconciliation Automation (5-File Architecture)

A Python automation system that produces one final Excel report (`output/Final_Report_<YYYY-MM-DD>.xlsx`) by integrating and reconciling **five auto-detected CSV/Excel files**.

---

## Input Files (Auto-Detected from `/input`)

Drop your files into the `input/` folder; the automation auto-detects them by name pattern:

| # | Input File | Auto-Detection Pattern | Key Columns Used |
|---|---|---|---|
| 1 | **Seller Orders Report** | Filename contains `Seller_Orders_Report` | `order release id`, `seller sku code` |
| 2 | **Order Flow Report** | Filename starts with `part-` | `sale_order_code`, `return_type` |
| 3 | **Prepaid Settlement** | Filename contains `prepaid` (and NOT `postpaid`) | `order release id`, `Settled_Amount` |
| 4 | **Postpaid Settlement** | Filename contains `postpaid` | `order release id`, `Settled_Amount` |
| 5 | **Master SKU File** | Filename contains `master` (CSV or XLSX) | `SKU Name`, `Style Name`, `Payout` |

> **Configuration Overrides:** You can explicitly configure filenames or column names in the `CONFIG_FILE_OVERRIDES` and `CONFIG_COLUMN_NAMES` dictionaries at the top of `run.py` or via CLI flags.

---

## Automation Logic & Business Rules

1. **Strict Text Reading & Numeric Conversion**:
   - All files are loaded with `dtype=str` and `keep_default_na=False`.
   - Strips whitespace and leading apostrophes (`'`) from key columns.
   - IDs are guaranteed to remain exact string representations (`data_type="s"`, format `"@"`).
   - Only `Settled_Amount` and `Payout` are converted to numeric float (strips currency symbols and commas; non-numeric $\rightarrow$ `0.0` with a warning).

2. **Scientific Notation Guard**:
   - Before processing, inspects `sale_order_code` in the Order Flow report.
   - If any value contains `"E+"` or matches scientific notation (e.g. `1.00282E+11`), **STOPS execution immediately** with clear guidance:
     ```text
     sale_order_code is in scientific notation. Re-export the Order Flow report without opening it in Excel, or format the column as Text.
     ```
   - Verifies that `sale_order_code` values are unique-looking (alerts if uniqueness ratio is far lower than row count).

3. **Step 1 — Base Table (Order Flow + SKU + Return Type)**:
   - Order Flow is the base table; **every single row is preserved**.
   - Left join: Order Flow `sale_order_code` == Orders `order release id` to fetch `seller sku code`.
   - If no match, populated as `"NOT FOUND"`. Rows are never dropped.
   - If `order release id` is duplicated in Orders, keeps first occurrence and logs a warning.
   - If the same `sale_order_code` appears on several Order Flow rows, keeps every row (each gets its SKU lookup) and logs repeated counts.
   - Maps `return_type` case-insensitively using a configurable mapping dictionary:
     * blank / empty / NaN $\rightarrow$ `"Delivered"`
     * `release_rto_refund` $\rightarrow$ `"RTO"`
     * `return_refund` $\rightarrow$ `"RVP"`
     * `exchange` $\rightarrow$ `"Exchange"`
     * Any other value $\rightarrow$ `"UNMAPPED: <raw value>"` (logs warning)

4. **Step 2 — Settlement Data (Prepaid + Postpaid Stack)**:
   - Prepaid file: extracts `order release id`, `Settled_Amount`, sets `Payment_Type = "Prepaid"`.
   - Postpaid file: extracts `order release id`, `Settled_Amount`, sets `Payment_Type = "Postpaid"`.
   - Stacks Postpaid below Prepaid into one unified table:
     `order release id | Settled_Amount | Payment_Type`
   - Keeps every row as it appears in the source files. Adds a **Grand Total** row at the bottom for `Settled_Amount`.

5. **Step 3 — Settlement Pivot & Merge**:
   - Aggregates `Settled_Amount` by `order release id` across the stacked table.
   - Left joins that total into the base table matching Order Flow `sale_order_code == order release id`. If an order release id has no settlement, sets `Settled_Amount = 0.0`.

6. **Step 4 — Master SKU Lookup (Style Name + Payout)**:
   - Left joins base table `seller sku code` == Master `SKU Name` (matching after whitespace trimming and ignoring upper/lower case).
   - Fetches `Style Name` and `Payout`.
   - If no match (including rows where `seller sku code` is `"NOT FOUND"`):
     `Style Name = "NOT FOUND"`, `Payout = 0.0`.
   - If `SKU Name` is duplicated in the master, keeps the first occurrence and logs a warning.
   - Collects the de-duplicated list of seller SKU codes not found in the master for the validation audit log.

---

## Output Excel Workbook (`/output/Final_Report_<YYYY-MM-DD>.xlsx`)

The workbook contains the **`Final` Sheet** matching the sample reference specification:

### **`Final` Sheet Columns & Formulas**:
EXACTLY these 11 columns in this exact sequence:
```text
order release id | seller sku code | Style Name | return_type | QTY | RETURN QTY | SALE QTY  | Payout | NET PRODUCT COST  | Settled_Amount | PROFIT
```

1. `order release id`: the Order Flow `sale_order_code` integer/text value.
2. `seller sku code`: fetched SKU code (highlighted in soft red if `"NOT FOUND"`).
3. `Style Name`: fetched Style Name from master (highlighted in soft red if `"NOT FOUND"`).
4. `return_type`: mapped status (`Delivered`, `RTO`, `RVP`, `Exchange`, or `UNMAPPED:`).
5. `QTY`: integer `1` per unit item line.
6. `RETURN QTY`: integer `0` for `Delivered`, `1` for `RTO`, `RVP`, `Exchange`.
7. `SALE QTY `: dynamic Excel formula `=E{row}-F{row}` (`QTY - RETURN QTY`).
8. `Payout`: numeric payout amount from master, formatted as currency (`#,##0.00`).
9. `NET PRODUCT COST `: dynamic Excel formula `=H{row}*G{row}` (`Payout * SALE QTY`).
10. `Settled_Amount`: total settlement looked up from settlement data, formatted as currency (`#,##0.00`).
11. `PROFIT`: dynamic Excel formula `=J{row}-I{row}` (`Settled_Amount - NET PRODUCT COST`).

> **Optional Settlement Sheets**: By default, the output generates only the clean `Final` sheet. If you wish to also include the unaggregated stacked settlement and pivot sheets (`Settlement_Data` & `Settlement_Pivot`), pass the `--include-settlement-sheets` CLI flag.

---

## Validation & Audit (Console & `/logs/log.txt` Only)

Validation and audit statistics are printed to the console and written to `logs/log.txt` (never cluttering the output workbook):
- Row counts for all five source files.
- Match rate and count of `"NOT FOUND"` SKUs.
- Count and de-duplicated list of SKUs missing from the master file.
- Count of repeated `sale_order_code` values and affected rows.
- Row breakdown per `return_type` and unmapped values.
- Summary of settlement totals (Prepaid, Postpaid, Combined, and Final report sum).
- Sum of `Payout` in Final report.
- **Reconciliation Audit**: Sum of `Settled_Amount` in `Final` vs Combined Settlement Total. If the final sum is lower, lists every `order release id` that was settled in Prepaid/Postpaid but did not match any order in Order Flow.

---

## Quick Start

### Installation
```bash
pip install -r requirements.txt
```

### CLI Execution
```bash
python run.py
```

### Web Studio Execution
```bash
python web_server.py
```
Open [http://localhost:8000](http://localhost:8000) in your browser.

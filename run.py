#!/usr/bin/env python3
"""
Order Flow & Settlement Reconciliation Automation
==================================================
Produces one final Excel workbook (/output/Final_Report_<YYYY-MM-DD>.xlsx)
from five auto-detected CSV/Excel reports:
  1. Seller Orders Report   -> filename contains "Seller_Orders_Report"
     Columns used: "order release id", "seller sku code"
  2. Order Flow Report      -> filename starts with "part-"
     Columns used: "sale_order_code", "return_type"
  3. Prepaid Settlement     -> filename contains "prepaid" but NOT "postpaid"
     Columns used: "order release id", "Settled_Amount"
  4. Postpaid Settlement    -> filename contains "postpaid"
     Columns used: "order release id", "Settled_Amount"
  5. Master SKU file        -> filename contains "master" (CSV or XLSX)
     Columns used: "SKU Name", "Style Name", "Payout"

Workbook Sheets:
  - "Final": EXACTLY 6 columns:
      order release id | seller sku code | Style Name | return_type | Settled_Amount | Payout
      ("order release id" = the Order Flow sale_order_code value, renamed)
      Grand-total row for Settled_Amount and Payout at the bottom.
  - "Settlement_Data": Stacked table of prepaid + postpaid records:
      order release id | Settled_Amount | Payment_Type
      with a grand-total row for Settled_Amount at the bottom.
  - "Settlement_Pivot": Aggregated pivot table by order release id:
      order release id | Prepaid_Settled | Postpaid_Settled | Total_Settled
      with a grand-total row at the bottom.

Scientific Notation Guard:
  - Halts execution with a clear error message if sale_order_code in Order Flow
    contains scientific notation pattern (e.g., '1.00282E+11' containing 'E+').
  - Validates uniqueness ratio of sale_order_code.
"""

import os
import sys
import glob
import re
import argparse
from datetime import datetime
from typing import Optional, Tuple, Dict, Any, List

# Ensure safe console output encoding on Windows terminals
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURATION SECTION (Overrides & Business Rules)
# Put file names below as overrides in case auto-detect is ambiguous.
# Column names also go in the config so they are easy to change.
# ==============================================================================
CONFIG_FILE_OVERRIDES = {
    "seller_orders_file": None,       # e.g., "input/ZbLNUV5X_Seller_Orders_Report.csv"
    "order_flow_file": None,          # e.g., "input/part-00824-...csv"
    "prepaid_settlement_file": None,  # e.g., "input/PREPAID DATA.csv"
    "postpaid_settlement_file": None, # e.g., "input/POSTPAID DATA.csv"
    "master_sku_file": None,          # e.g., "input/MASTER_SHEET_DATA.xlsx"
}

CONFIG_COLUMN_NAMES = {
    "seller_orders": {
        "order_release_id": "order release id",
        "seller_sku_code": "seller sku code",
    },
    "order_flow": {
        "sale_order_code": "sale_order_code",
        "return_type": "return_type",
    },
    "prepaid_settlement": {
        "order_release_id": "order release id",
        "settled_amount": "Settled_Amount",
    },
    "postpaid_settlement": {
        "order_release_id": "order release id",
        "settled_amount": "Settled_Amount",
    },
    "master_sku": {
        "sku_name": "SKU Name",
        "style_name": "Style Name",
        "payout": "Payout",
    },
}

# Return type to status mapping dictionary
RETURN_TYPE_STATUS_MAPPING = {
    "release_rto_refund": "RTO",
    "return_refund": "RVP",
    "exchange": "Exchange",
}
DEFAULT_BLANK_STATUS = "Delivered"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT_DIR = os.path.join(BASE_DIR, "input")
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DEFAULT_LOG_PATH = os.path.join(BASE_DIR, "logs", "log.txt")


class DualLogger:
    """Logs messages simultaneously to the console and to log.txt."""
    def __init__(self, log_path: str):
        self.log_path = log_path
        self.entries = []
        os.makedirs(os.path.dirname(os.path.abspath(log_path)), exist_ok=True)
        with open(self.log_path, "w", encoding="utf-8") as f:
            f.write(f"=== ORDER FLOW & SETTLEMENT PIPELINE AUDIT LOG: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n\n")

    def log(self, message: str = ""):
        print(message)
        self.entries.append(message)
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(message + "\n")

    def warning(self, message: str):
        self.log(f"[WARNING] {message}")

    def error(self, message: str):
        self.log(f"[ERROR] {message}")

    def info(self, message: str):
        self.log(message)


def clean_key_series(series: pd.Series) -> pd.Series:
    """
    Strips leading & trailing whitespace and removes leading apostrophe (').
    Guarantees pure string representation for safe ID and SKU lookups.
    """
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.lstrip("'")
        .str.strip()
    )


def clean_amount_series(series: pd.Series, label: str, logger: DualLogger) -> pd.Series:
    """
    Converts amount/payout column to numeric float safely:
    - Strips commas, currency symbols ($, Rs, INR), and whitespace.
    - Converts non-numeric / invalid entries to 0.0 with a logged warning.
    """
    cleaned = []
    warning_count = 0
    sample_bad = []

    for val in series:
        if val is None or pd.isna(val):
            cleaned.append(0.0)
            continue

        raw = str(val).strip()
        if not raw:
            cleaned.append(0.0)
            continue

        sanitized = re.sub(r"[,\$₹€£]|Rs\.?|INR", "", raw, flags=re.IGNORECASE).strip()
        try:
            num = float(sanitized)
            cleaned.append(num)
        except ValueError:
            cleaned.append(0.0)
            warning_count += 1
            if len(sample_bad) < 5:
                sample_bad.append(raw)

    if warning_count > 0:
        logger.warning(
            f"Found {warning_count} non-numeric value(s) in {label}. Converted to 0.0. "
            f"Samples: {sample_bad}"
        )

    return pd.Series(cleaned, index=series.index, dtype=float)


def check_scientific_notation_guard(series: pd.Series, col_name: str, logger: DualLogger):
    """
    SCIENTIFIC NOTATION GUARD:
    Before processing, check 'sale_order_code' in the Order Flow report.
    If any value matches the pattern like 1.00282E+11 (contains 'E+'),
    STOP and print a clear error:
    "sale_order_code is in scientific notation. Re-export the Order Flow report without opening it in Excel, or format the column as Text."
    Do not continue with truncated IDs.
    """
    raw_strings = series.fillna("").astype(str).str.strip()
    sci_mask = raw_strings.str.contains(r"[Ee]\+", regex=True)

    if sci_mask.any():
        bad_samples = raw_strings[sci_mask].head(3).tolist()
        err_msg = (
            f"sale_order_code is in scientific notation. "
            f"Re-export the Order Flow report without opening it in Excel, or format the column as Text."
        )
        logger.error(f"{err_msg} (Found {sci_mask.sum()} corrupted rows, samples: {bad_samples})")
        raise ValueError(err_msg)

    # Check that sale_order_code values are unique-looking
    unique_count = raw_strings.nunique()
    total_count = len(raw_strings)
    if total_count > 0:
        ratio = unique_count / total_count
        if ratio < 0.5:
            logger.warning(
                f"UNIQUENESS ALERT: '{col_name}' has only {unique_count:,} unique values across {total_count:,} rows "
                f"({ratio * 100:.1f}% uniqueness ratio). Number of unique values is far lower than row count."
            )


def map_return_type_to_status(val: str, mapping: dict = None, blank_status: str = None) -> str:
    """
    Maps return_type to final status case-insensitively.
    Unknown values are mapped to 'UNMAPPED: <raw value>'.
    """
    if mapping is None:
        mapping = RETURN_TYPE_STATUS_MAPPING
    if blank_status is None:
        blank_status = DEFAULT_BLANK_STATUS

    raw = str(val).strip() if val is not None else ""
    if not raw or raw.lower() in ("nan", "none", "null"):
        return blank_status

    normalized = raw.lower()
    for map_key, map_status in mapping.items():
        if normalized == map_key.strip().lower():
            return map_status

    return f"UNMAPPED: {raw}"


def find_column_name(
    df: pd.DataFrame,
    primary_config_name: str,
    possible_names: List[str],
    file_label: str,
    filename: str
) -> str:
    """
    Finds a matching column name in DataFrame:
    1. Checks configured primary name first (case-insensitively).
    2. Then checks possible candidate aliases.
    Raises descriptive KeyError if none of the candidates match.
    """
    col_map = {str(c).strip().lower(): str(c).strip() for c in df.columns}
    
    # 1. Primary config
    if primary_config_name and primary_config_name.strip().lower() in col_map:
        return col_map[primary_config_name.strip().lower()]

    # 2. Candidate aliases
    for candidate in possible_names:
        cand_norm = candidate.strip().lower()
        if cand_norm in col_map:
            return col_map[cand_norm]

    existing = list(df.columns)
    all_candidates = [primary_config_name] + [p for p in possible_names if p != primary_config_name]
    raise KeyError(
        f"Missing required column in {file_label} ('{os.path.basename(filename)}').\n"
        f"Expected one of: {all_candidates}\n"
        f"Columns present in file: {existing}"
    )


def auto_detect_input_files(
    input_dir: str,
    logger: DualLogger,
    overrides: Dict[str, Optional[str]] = None
) -> Tuple[str, str, str, str, str]:
    """
    Auto-detects the 5 required files from input_dir:
      1. Seller Orders: contains "Seller_Orders_Report"
      2. Order Flow: starts with "part-"
      3. Prepaid Settlement: contains "prepaid" but NOT "postpaid"
      4. Postpaid Settlement: contains "postpaid"
      5. Master SKU: contains "master" (CSV or XLSX)
    Respects overrides if specified.
    """
    if not os.path.exists(input_dir):
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    all_files = glob.glob(os.path.join(input_dir, "*"))
    supported_files = [
        f for f in all_files
        if f.lower().endswith((".csv", ".xlsx", ".xls"))
    ]
    if not supported_files:
        raise FileNotFoundError(f"No CSV or Excel files found in input directory: {input_dir}")

    overrides = overrides or {}

    # 1. Seller Orders
    orders_path = overrides.get("seller_orders_file")
    if not orders_path or not os.path.exists(orders_path):
        matches = [f for f in supported_files if "seller_orders_report" in os.path.basename(f).lower() or ("seller" in os.path.basename(f).lower() and "orders" in os.path.basename(f).lower())]
        if not matches:
            raise FileNotFoundError(
                f"Missing Seller Orders file in '{input_dir}'. Expected filename containing 'Seller_Orders_Report'."
            )
        matches.sort(key=os.path.getmtime, reverse=True)
        orders_path = matches[0]

    # 2. Order Flow
    flow_path = overrides.get("order_flow_file")
    if not flow_path or not os.path.exists(flow_path):
        matches = [f for f in supported_files if os.path.basename(f).lower().startswith("part-") or "flow" in os.path.basename(f).lower()]
        if not matches:
            raise FileNotFoundError(
                f"Missing Order Flow file in '{input_dir}'. Expected filename starting with 'part-'."
            )
        matches.sort(key=os.path.getmtime, reverse=True)
        flow_path = matches[0]

    # 3. Prepaid Settlement
    prepaid_path = overrides.get("prepaid_settlement_file")
    if not prepaid_path or not os.path.exists(prepaid_path):
        matches = [
            f for f in supported_files
            if "prepaid" in os.path.basename(f).lower() and "postpaid" not in os.path.basename(f).lower()
        ]
        if not matches:
            raise FileNotFoundError(
                f"Missing Prepaid Settlement file in '{input_dir}'. Expected filename containing 'prepaid' (and not 'postpaid')."
            )
        matches.sort(key=os.path.getmtime, reverse=True)
        prepaid_path = matches[0]

    # 4. Postpaid Settlement
    postpaid_path = overrides.get("postpaid_settlement_file")
    if not postpaid_path or not os.path.exists(postpaid_path):
        matches = [f for f in supported_files if "postpaid" in os.path.basename(f).lower()]
        if not matches:
            raise FileNotFoundError(
                f"Missing Postpaid Settlement file in '{input_dir}'. Expected filename containing 'postpaid'."
            )
        matches.sort(key=os.path.getmtime, reverse=True)
        postpaid_path = matches[0]

    # 5. Master SKU
    master_path = overrides.get("master_sku_file")
    if not master_path or not os.path.exists(master_path):
        matches = [f for f in supported_files if "master" in os.path.basename(f).lower()]
        if not matches:
            raise FileNotFoundError(
                f"Missing Master SKU file in '{input_dir}'. Expected filename containing 'master' (.csv or .xlsx)."
            )
        matches.sort(key=os.path.getmtime, reverse=True)
        master_path = matches[0]

    logger.info(f"Detected Seller Orders Report   : {os.path.basename(orders_path)}")
    logger.info(f"Detected Order Flow Report     : {os.path.basename(flow_path)}")
    logger.info(f"Detected Prepaid Settlement    : {os.path.basename(prepaid_path)}")
    logger.info(f"Detected Postpaid Settlement   : {os.path.basename(postpaid_path)}")
    logger.info(f"Detected Master SKU File       : {os.path.basename(master_path)}")

    return orders_path, flow_path, prepaid_path, postpaid_path, master_path


find_input_files = auto_detect_input_files


def load_file_as_text(filepath: str, label: str, logger: DualLogger) -> pd.DataFrame:
    """
    Reads a CSV or Excel file strictly with dtype=str and keep_default_na=False across all columns.
    Prevents floating-point conversions, loss of precision, and scientific notation.
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".xlsx", ".xls"):
        try:
            df = pd.read_excel(filepath, dtype=str, keep_default_na=False)
            logger.info(f"Loaded {label} from Excel '{os.path.basename(filepath)}' ({len(df):,} rows).")
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            raise RuntimeError(f"Error reading Excel {label} file '{filepath}': {str(e)}")

    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    df = None
    last_err = None

    for enc in encodings:
        try:
            df = pd.read_csv(
                filepath,
                dtype=str,
                keep_default_na=False,
                encoding=enc,
                low_memory=False
            )
            logger.info(f"Loaded {label} from CSV '{os.path.basename(filepath)}' using encoding '{enc}' ({len(df):,} rows).")
            break
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except Exception as e:
            raise RuntimeError(f"Error reading {label} file '{filepath}': {str(e)}")

    if df is None:
        raise RuntimeError(f"Failed to read {label} file '{filepath}'. Encoding error: {last_err}")

    # Strip whitespace from all header names
    df.columns = [str(c).strip() for c in df.columns]
    return df


load_csv_as_text = load_file_as_text


def write_excel_final_report(
    xlsx_path: str,
    final_df: pd.DataFrame,
    settlement_data_df: pd.DataFrame = None,
    settlement_pivot_df: pd.DataFrame = None,
    logger: DualLogger = None,
    include_settlement_sheets: bool = False
) -> str:
    """
    Creates the final Excel workbook matching the exact sample reference specification:
      Sheet 'Final': EXACTLY 11 columns:
        order release id | seller sku code | Style Name | return_type | QTY | RETURN QTY | SALE QTY  | Payout | NET PRODUCT COST  | Settled_Amount | PROFIT
      Formulas:
        SALE QTY         : =E{row}-F{row}
        NET PRODUCT COST : =H{row}*G{row}
        PROFIT           : =J{row}-I{row}
    """
    wb = openpyxl.Workbook()

    # Styling Palettes
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    
    data_font = Font(name="Calibri", size=10, color="000000")
    mono_font = Font(name="Calibri", size=10, color="000000")

    warn_notfound_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    warn_notfound_font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    
    warn_unmapped_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    warn_unmapped_font = Font(name="Calibri", size=10, bold=True, color="92400E")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )
    total_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="0F172A"),
        bottom=Side(style="double", color="0F172A")
    )

    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # =========================================================================
    # SHEET 1: "Final"
    # Columns: order release id | seller sku code | Style Name | return_type |
    #          QTY | RETURN QTY | SALE QTY  | Payout | NET PRODUCT COST  | Settled_Amount | PROFIT
    # =========================================================================
    ws_final = wb.active
    ws_final.title = "Final"
    ws_final.views.sheetView[0].showGridLines = True

    final_headers = [
        "order release id",
        "seller sku code",
        "Style Name",
        "return_type",
        "QTY",
        "RETURN QTY",
        "SALE QTY ",
        "Payout",
        "NET PRODUCT COST ",
        "Settled_Amount",
        "PROFIT"
    ]

    # Header Row
    for col_idx, col_name in enumerate(final_headers, start=1):
        cell = ws_final.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border
        cell.number_format = "@"

    # Data Rows
    for row_idx, row in enumerate(final_df.itertuples(index=False), start=2):
        # 1. order release id (numeric integer when digits, formatted as 0)
        raw_id = row[0]
        try:
            val_id = int(str(raw_id).strip())
            c1 = ws_final.cell(row=row_idx, column=1, value=val_id)
            c1.data_type = "n"
            c1.number_format = "0"
        except (ValueError, TypeError):
            c1 = ws_final.cell(row=row_idx, column=1, value=str(raw_id))
            c1.data_type = "s"
            c1.number_format = "@"
        c1.font = data_font
        c1.alignment = align_left
        c1.border = thin_border

        # 2. seller sku code (text)
        sku_val = str(row[1])
        c2 = ws_final.cell(row=row_idx, column=2, value=sku_val)
        c2.data_type = "s"
        c2.number_format = "@"
        c2.alignment = align_left
        c2.border = thin_border
        if sku_val == "NOT FOUND":
            c2.fill = warn_notfound_fill
            c2.font = warn_notfound_font
        else:
            c2.font = data_font

        # 3. Style Name (text)
        style_val = str(row[2])
        c3 = ws_final.cell(row=row_idx, column=3, value=style_val)
        c3.data_type = "s"
        c3.number_format = "@"
        c3.alignment = align_left
        c3.border = thin_border
        if style_val == "NOT FOUND":
            c3.fill = warn_notfound_fill
            c3.font = warn_notfound_font
        else:
            c3.font = data_font

        # 4. return_type (text, mapped value)
        ret_val = str(row[3])
        c4 = ws_final.cell(row=row_idx, column=4, value=ret_val)
        c4.data_type = "s"
        c4.number_format = "@"
        c4.alignment = align_center
        c4.border = thin_border
        if ret_val.startswith("UNMAPPED:"):
            c4.fill = warn_unmapped_fill
            c4.font = warn_unmapped_font
        else:
            c4.font = data_font

        # 5. QTY (numeric integer 1)
        qty_val = int(row[4]) if len(row) > 4 and row[4] is not None else 1
        c5 = ws_final.cell(row=row_idx, column=5, value=qty_val)
        c5.data_type = "n"
        c5.number_format = "General"
        c5.font = data_font
        c5.alignment = align_center
        c5.border = thin_border

        # 6. RETURN QTY (numeric integer: 0 for Delivered, 1 for Returned/RTO/RVP/Exchange)
        ret_qty = int(row[5]) if len(row) > 5 and row[5] is not None else (0 if str(ret_val).strip().lower() == "delivered" else 1)
        c6 = ws_final.cell(row=row_idx, column=6, value=ret_qty)
        c6.data_type = "n"
        c6.number_format = "General"
        c6.font = data_font
        c6.alignment = align_center
        c6.border = thin_border

        # 7. SALE QTY  (Formula =E{row_idx}-F{row_idx})
        c7 = ws_final.cell(row=row_idx, column=7, value=f"=E{row_idx}-F{row_idx}")
        c7.number_format = "General"
        c7.font = data_font
        c7.alignment = align_center
        c7.border = thin_border

        # 8. Payout (numeric float, Master SKU payout)
        payout_val = float(row[7]) if len(row) > 7 and row[7] is not None else 0.0
        c8 = ws_final.cell(row=row_idx, column=8, value=payout_val)
        c8.data_type = "n"
        c8.number_format = "#,##0.00"
        c8.font = data_font
        c8.alignment = align_right
        c8.border = thin_border

        # 9. NET PRODUCT COST  (Formula =H{row_idx}*G{row_idx})
        c9 = ws_final.cell(row=row_idx, column=9, value=f"=H{row_idx}*G{row_idx}")
        c9.number_format = "#,##0.00"
        c9.font = data_font
        c9.alignment = align_right
        c9.border = thin_border

        # 10. Settled_Amount (numeric float, from settlement)
        amt_val = float(row[9]) if len(row) > 9 and row[9] is not None else 0.0
        c10 = ws_final.cell(row=row_idx, column=10, value=amt_val)
        c10.data_type = "n"
        c10.number_format = "#,##0.00"
        c10.font = data_font
        c10.alignment = align_right
        c10.border = thin_border

        # 11. PROFIT (Formula =J{row_idx}-I{row_idx})
        c11 = ws_final.cell(row=row_idx, column=11, value=f"=J{row_idx}-I{row_idx}")
        c11.number_format = "#,##0.00"
        c11.font = Font(name="Calibri", size=11)

    # Set exact column widths matching sample reference
    col_widths = {
        "A": 20.0,
        "B": 27.0,
        "C": 19.0,
        "D": 18.0,
        "E": 13.0,
        "F": 13.0,
        "G": 13.0,
        "H": 13.0,
        "I": 19.140625,
        "J": 23.0,
        "K": 13.0
    }
    for col_letter, width in col_widths.items():
        ws_final.column_dimensions[col_letter].width = width

    # Optional sheets: only created if include_settlement_sheets is explicitly requested
    if include_settlement_sheets and settlement_data_df is not None and settlement_pivot_df is not None:
        # =========================================================================
        # SHEET 2: "Settlement_Data"
        # Columns: order release id | Settled_Amount | Payment_Type (with Grand Total)
        # =========================================================================
        ws_data = wb.create_sheet(title="Settlement_Data")
        ws_data.views.sheetView[0].showGridLines = True

        data_headers = ["order release id", "Settled_Amount", "Payment_Type"]
        for col_idx, col_name in enumerate(data_headers, start=1):
            cell = ws_data.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = thin_border
            cell.number_format = "@"

        total_settled_data = 0.0
        for row_idx, row in enumerate(settlement_data_df.itertuples(index=False), start=2):
            # order release id
            c1 = ws_data.cell(row=row_idx, column=1, value=str(row[0]))
            c1.data_type = "s"
            c1.number_format = "@"
            c1.font = mono_font
            c1.alignment = align_left
            c1.border = thin_border

            # Settled_Amount
            amt = float(row[1])
            total_settled_data += amt
            c2 = ws_data.cell(row=row_idx, column=2, value=amt)
            c2.data_type = "n"
            c2.number_format = "#,##0.00"
            c2.font = data_font
            c2.alignment = align_right
            c2.border = thin_border

            # Payment_Type
            c3 = ws_data.cell(row=row_idx, column=3, value=str(row[2]))
            c3.data_type = "s"
            c3.number_format = "@"
            c3.font = data_font
            c3.alignment = align_center
            c3.border = thin_border

        # Grand Total Row at bottom of "Settlement_Data"
        last_row_data = ws_data.max_row + 1
        t1 = ws_data.cell(row=last_row_data, column=1, value="Grand Total")
        t1.font = total_font
        t1.fill = total_fill
        t1.alignment = align_left
        t1.border = total_border

        t2 = ws_data.cell(row=last_row_data, column=2, value=round(total_settled_data, 2))
        t2.data_type = "n"
        t2.number_format = "#,##0.00"
        t2.font = total_font
        t2.fill = total_fill
        t2.alignment = align_right
        t2.border = total_border

        t3 = ws_data.cell(row=last_row_data, column=3, value="")
        t3.fill = total_fill
        t3.border = total_border

        for col_idx in range(1, 4):
            col_letter = get_column_letter(col_idx)
            ws_data.column_dimensions[col_letter].width = 22

        # =========================================================================
        # SHEET 3: "Settlement_Pivot"
        # Columns: order release id | Prepaid_Settled | Postpaid_Settled | Total_Settled
        # =========================================================================
        ws_pivot = wb.create_sheet(title="Settlement_Pivot")
        ws_pivot.views.sheetView[0].showGridLines = True

        pivot_headers = ["order release id", "Prepaid_Settled", "Postpaid_Settled", "Total_Settled"]
        for col_idx, col_name in enumerate(pivot_headers, start=1):
            cell = ws_pivot.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = thin_border
            cell.number_format = "@"

        sum_prepaid = 0.0
        sum_postpaid = 0.0
        sum_total = 0.0

        for row_idx, row in enumerate(settlement_pivot_df.itertuples(index=False), start=2):
            # order release id
            c1 = ws_pivot.cell(row=row_idx, column=1, value=str(row[0]))
            c1.data_type = "s"
            c1.number_format = "@"
            c1.font = mono_font
            c1.alignment = align_left
            c1.border = thin_border

            # Prepaid_Settled
            v_pre = float(row[1])
            sum_prepaid += v_pre
            c2 = ws_pivot.cell(row=row_idx, column=2, value=v_pre)
            c2.data_type = "n"
            c2.number_format = "#,##0.00"
            c2.font = data_font
            c2.alignment = align_right
            c2.border = thin_border

            # Postpaid_Settled
            v_post = float(row[2])
            sum_postpaid += v_post
            c3 = ws_pivot.cell(row=row_idx, column=3, value=v_post)
            c3.data_type = "n"
            c3.number_format = "#,##0.00"
            c3.font = data_font
            c3.alignment = align_right
            c3.border = thin_border

            # Total_Settled
            v_tot = float(row[3])
            sum_total += v_tot
            c4 = ws_pivot.cell(row=row_idx, column=4, value=v_tot)
            c4.data_type = "n"
            c4.number_format = "#,##0.00"
            c4.font = data_font
            c4.alignment = align_right
            c4.border = thin_border

        # Grand Total Row at bottom of "Settlement_Pivot"
        last_row_piv = ws_pivot.max_row + 1
        pt1 = ws_pivot.cell(row=last_row_piv, column=1, value="Grand Total")
        pt1.font = total_font
        pt1.fill = total_fill
        pt1.alignment = align_left
        pt1.border = total_border

        pt2 = ws_pivot.cell(row=last_row_piv, column=2, value=round(sum_prepaid, 2))
        pt2.data_type = "n"
        pt2.number_format = "#,##0.00"
        pt2.font = total_font
        pt2.fill = total_fill
        pt2.alignment = align_right
        pt2.border = total_border

        pt3 = ws_pivot.cell(row=last_row_piv, column=3, value=round(sum_postpaid, 2))
        pt3.data_type = "n"
        pt3.number_format = "#,##0.00"
        pt3.font = total_font
        pt3.fill = total_fill
        pt3.alignment = align_right
        pt3.border = total_border

        pt4 = ws_pivot.cell(row=last_row_piv, column=4, value=round(sum_total, 2))
        pt4.data_type = "n"
        pt4.number_format = "#,##0.00"
        pt4.font = total_font
        pt4.fill = total_fill
        pt4.alignment = align_right
        pt4.border = total_border

        for col_idx in range(1, 5):
            col_letter = get_column_letter(col_idx)
            ws_pivot.column_dimensions[col_letter].width = 22

    # Save complete workbook
    try:
        wb.save(xlsx_path)
        return xlsx_path
    except PermissionError:
        timestamp_suffix = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        alt_filename = f"Final_Report_{timestamp_suffix}.xlsx"
        alt_xlsx_path = os.path.join(os.path.dirname(xlsx_path), alt_filename)
        if logger:
            logger.warning(
                f"File '{xlsx_path}' is currently open in another program (e.g. Excel). "
                f"Successfully saved to '{alt_xlsx_path}' instead."
            )
        wb.save(alt_xlsx_path)
        return alt_xlsx_path


def process_order_flow_pipeline(
    orders_path: str = None,
    flow_path: str = None,
    prepaid_path: str = None,
    postpaid_path: str = None,
    master_path: str = None,
    input_dir: str = DEFAULT_INPUT_DIR,
    output_dir: str = DEFAULT_OUTPUT_DIR,
    log_path: str = DEFAULT_LOG_PATH,
    mapping_dict: dict = None,
    blank_status: str = None,
    include_settlement_sheets: bool = False
) -> Dict[str, Any]:
    """
    Executes the complete 5-file pipeline:
      - Reads all 5 files strictly as text (pandas dtype=str).
      - Scientific Notation Guard on Order Flow 'sale_order_code'.
      - Step 1: Base table left join on sale_order_code == order release id + return status mapping.
      - Step 2: Stack prepaid + postpaid settlement files with Grand Total.
      - Step 3: Pivot total by order release id and left join into final table.
      - Step 4: Master SKU lookup for Style Name and numeric Payout (case-insensitive, whitespace-trimmed).
      - Output: /output/Final_Report_<YYYY-MM-DD>.xlsx with sheets Final, Settlement_Data, Settlement_Pivot.
      - Validation & Audit: Detailed reconciliation and unlinked settlement checks logged to console + log.txt.
    """
    logger = DualLogger(log_path)
    logger.info("=" * 80)
    logger.info("ORDER FLOW & SETTLEMENT RECONCILIATION AUTOMATION (5-FILE PIPELINE)")
    logger.info("=" * 80)
    logger.info(f"Timestamp : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 1. Resolve / Auto-Detect Files
    overrides = dict(CONFIG_FILE_OVERRIDES)
    if orders_path: overrides["seller_orders_file"] = orders_path
    if flow_path: overrides["order_flow_file"] = flow_path
    if prepaid_path: overrides["prepaid_settlement_file"] = prepaid_path
    if postpaid_path: overrides["postpaid_settlement_file"] = postpaid_path
    if master_path: overrides["master_sku_file"] = master_path

    detected_orders, detected_flow, detected_prepaid, detected_postpaid, detected_master = auto_detect_input_files(
        input_dir=input_dir,
        logger=logger,
        overrides=overrides
    )

    logger.info("-" * 80)

    # 2. Read All Five Files Strictly as Text
    orders_df = load_file_as_text(detected_orders, "Seller Orders Report", logger)
    flow_df = load_file_as_text(detected_flow, "Order Flow Report", logger)
    prepaid_df = load_file_as_text(detected_prepaid, "Prepaid Settlement Report", logger)
    postpaid_df = load_file_as_text(detected_postpaid, "Postpaid Settlement Report", logger)
    master_df = load_file_as_text(detected_master, "Master SKU Report", logger)

    orders_count = len(orders_df)
    flow_count = len(flow_df)
    prepaid_count = len(prepaid_df)
    postpaid_count = len(postpaid_df)
    master_count = len(master_df)

    # Validate Column Presence using Config + Aliases
    cols_cfg = CONFIG_COLUMN_NAMES

    orders_release_id_col = find_column_name(
        orders_df,
        cols_cfg["seller_orders"]["order_release_id"],
        ["order release id", "order_release_id"],
        "Seller Orders",
        detected_orders
    )
    orders_sku_col = find_column_name(
        orders_df,
        cols_cfg["seller_orders"]["seller_sku_code"],
        ["seller sku code", "seller_sku_code", "seller sku on channel", "sku code", "sku"],
        "Seller Orders",
        detected_orders
    )

    flow_sale_code_col = find_column_name(
        flow_df,
        cols_cfg["order_flow"]["sale_order_code"],
        ["sale_order_code", "sale order code"],
        "Order Flow",
        detected_flow
    )
    flow_return_type_col = find_column_name(
        flow_df,
        cols_cfg["order_flow"]["return_type"],
        ["return_type", "return type"],
        "Order Flow",
        detected_flow
    )

    prepaid_release_col = find_column_name(
        prepaid_df,
        cols_cfg["prepaid_settlement"]["order_release_id"],
        ["order release id", "order_release_id"],
        "Prepaid Settlement",
        detected_prepaid
    )
    prepaid_amount_col = find_column_name(
        prepaid_df,
        cols_cfg["prepaid_settlement"]["settled_amount"],
        ["Settled_Amount", "settled_amount", "Settled Amount"],
        "Prepaid Settlement",
        detected_prepaid
    )

    postpaid_release_col = find_column_name(
        postpaid_df,
        cols_cfg["postpaid_settlement"]["order_release_id"],
        ["order release id", "order_release_id"],
        "Postpaid Settlement",
        detected_postpaid
    )
    postpaid_amount_col = find_column_name(
        postpaid_df,
        cols_cfg["postpaid_settlement"]["settled_amount"],
        ["Settled_Amount", "settled_amount", "Settled Amount"],
        "Postpaid Settlement",
        detected_postpaid
    )

    # Master SKU columns
    master_sku_col = find_column_name(
        master_df,
        cols_cfg["master_sku"]["sku_name"],
        ["sku name", "seller sku on channel", "sku code", "sku_name", "sku", "skuname"],
        "Master SKU",
        detected_master
    )
    master_style_col = find_column_name(
        master_df,
        cols_cfg["master_sku"]["style_name"],
        ["style name", "style code", "style_name", "style", "stylename"],
        "Master SKU",
        detected_master
    )
    master_payout_col = find_column_name(
        master_df,
        cols_cfg["master_sku"]["payout"],
        ["payout", "pay out", "pay_out", "production cost", "payout amount"],
        "Master SKU",
        detected_master
    )

    # -------------------------------------------------------------------------
    # SCIENTIFIC NOTATION GUARD
    # -------------------------------------------------------------------------
    logger.info(f"\n--- Checking Scientific Notation Guard on Order Flow '{flow_sale_code_col}' ---")
    check_scientific_notation_guard(flow_df[flow_sale_code_col], flow_sale_code_col, logger)
    logger.info(f"[PASS] Scientific notation guard passed cleanly for '{flow_sale_code_col}'.")

    # -------------------------------------------------------------------------
    # STEP 1 - BASE TABLE (Order Flow + SKU + return type)
    # -------------------------------------------------------------------------
    logger.info("\n--- STEP 1: Building Base Table (Order Flow + Seller SKU + Return Status) ---")
    orders_df["_clean_order_release_id"] = clean_key_series(orders_df[orders_release_id_col])
    orders_df["_clean_seller_sku"] = orders_df[orders_sku_col].fillna("").astype(str).str.strip().str.lstrip("'").str.strip()

    flow_df["_clean_sale_order_code"] = clean_key_series(flow_df[flow_sale_code_col])
    flow_df["_row_order"] = range(len(flow_df))

    # Check repeated sale_order_code in Order Flow
    flow_counts = flow_df["_clean_sale_order_code"].value_counts()
    repeated_codes = flow_counts[flow_counts > 1]
    repeated_count = len(repeated_codes)
    total_repeated_rows = int(repeated_codes.sum()) if repeated_count > 0 else 0
    if repeated_count > 0:
        logger.info(
            f"Order Flow contains {repeated_count} repeated '{flow_sale_code_col}' value(s) across {total_repeated_rows} rows. "
            f"All rows are kept, each receiving its SKU lookup."
        )

    # Handle duplicates in Seller Orders: keep first, log warning
    dup_mask = orders_df.duplicated(subset=["_clean_order_release_id"], keep=False)
    dup_count = int(dup_mask.sum())
    if dup_count > 0:
        unique_dups = orders_df.loc[dup_mask, "_clean_order_release_id"].nunique()
        logger.warning(
            f"Seller Orders report contains {dup_count} duplicate rows ({unique_dups} unique '{orders_release_id_col}' keys). "
            f"Keeping the first occurrence and ignoring subsequent duplicates."
        )
        sample_dups = orders_df.loc[dup_mask, "_clean_order_release_id"].unique()[:5]
        logger.warning(f"  -> Sample duplicate keys: {list(sample_dups)}")

    orders_dedup = orders_df.drop_duplicates(subset=["_clean_order_release_id"], keep="first")

    # Left join to preserve every row in Order Flow:
    # Order Flow "sale_order_code" == Orders "order release id"
    lookup_table = orders_dedup[["_clean_order_release_id", "_clean_seller_sku"]]
    base_merged = flow_df.merge(
        lookup_table,
        left_on="_clean_sale_order_code",
        right_on="_clean_order_release_id",
        how="left"
    )
    base_merged.sort_values(by="_row_order", inplace=True)

    # Missing matches -> "NOT FOUND"
    base_merged["seller sku code"] = base_merged["_clean_seller_sku"].fillna("NOT FOUND")
    base_merged.loc[base_merged["seller sku code"].str.strip() == "", "seller sku code"] = "NOT FOUND"

    matched_sku_count = int((base_merged["seller sku code"] != "NOT FOUND").sum())
    not_found_sku_count = flow_count - matched_sku_count

    # Map return_type
    base_merged["mapped_return_type"] = base_merged[flow_return_type_col].apply(
        lambda v: map_return_type_to_status(v, mapping=mapping_dict, blank_status=blank_status)
    )

    unmapped_mask = base_merged["mapped_return_type"].str.startswith("UNMAPPED: ")
    unmapped_count = int(unmapped_mask.sum())
    unmapped_samples = base_merged.loc[unmapped_mask, flow_return_type_col].unique().tolist() if unmapped_count > 0 else []
    if unmapped_count > 0:
        logger.warning(f"Found {unmapped_count} row(s) with UNMAPPED return_type values: {unmapped_samples}")

    # -------------------------------------------------------------------------
    # STEP 2 - SETTLEMENT DATA (PREPAID + POSTPAID IN ONE SHEET)
    # -------------------------------------------------------------------------
    logger.info("\n--- STEP 2: Processing Settlement Data (Prepaid + Postpaid Stack) ---")

    # Clean Prepaid
    prepaid_clean = pd.DataFrame({
        "order release id": clean_key_series(prepaid_df[prepaid_release_col]),
        "Settled_Amount": clean_amount_series(prepaid_df[prepaid_amount_col], "Prepaid", logger),
        "Payment_Type": "Prepaid"
    })

    # Clean Postpaid
    postpaid_clean = pd.DataFrame({
        "order release id": clean_key_series(postpaid_df[postpaid_release_col]),
        "Settled_Amount": clean_amount_series(postpaid_df[postpaid_amount_col], "Postpaid", logger),
        "Payment_Type": "Postpaid"
    })

    # Stack: append postpaid rows below prepaid rows
    settlement_stacked_df = pd.concat([prepaid_clean, postpaid_clean], ignore_index=True)

    total_prepaid_amount = float(prepaid_clean["Settled_Amount"].sum())
    total_postpaid_amount = float(postpaid_clean["Settled_Amount"].sum())
    total_combined_settlement = float(settlement_stacked_df["Settled_Amount"].sum())

    # -------------------------------------------------------------------------
    # STEP 3 - SETTLED AMOUNT INTO FINAL SHEET
    # -------------------------------------------------------------------------
    logger.info("\n--- STEP 3: Aggregating Settlement Pivot & Merging into Final Sheet ---")

    # Pivot / Sum by order release id across stacked table
    pivot_agg = settlement_stacked_df.pivot_table(
        index="order release id",
        columns="Payment_Type",
        values="Settled_Amount",
        aggfunc="sum",
        fill_value=0.0
    ).reset_index()

    # Ensure both columns exist
    if "Prepaid" not in pivot_agg.columns: pivot_agg["Prepaid"] = 0.0
    if "Postpaid" not in pivot_agg.columns: pivot_agg["Postpaid"] = 0.0

    pivot_agg["Total_Settled"] = pivot_agg["Prepaid"] + pivot_agg["Postpaid"]
    pivot_export_df = pd.DataFrame({
        "order release id": pivot_agg["order release id"],
        "Prepaid_Settled": pivot_agg["Prepaid"],
        "Postpaid_Settled": pivot_agg["Postpaid"],
        "Total_Settled": pivot_agg["Total_Settled"]
    })

    # Lookup total into base table matching Order Flow "sale_order_code" == "order release id"
    settlement_totals_lookup = pivot_export_df[["order release id", "Total_Settled"]].copy()
    base_with_settlement = base_merged.merge(
        settlement_totals_lookup,
        left_on="_clean_sale_order_code",
        right_on="order release id",
        how="left"
    )
    base_with_settlement.sort_values(by="_row_order", inplace=True)

    # Release ids with no settlement -> 0.0
    base_with_settlement["Settled_Amount"] = base_with_settlement["Total_Settled"].fillna(0.0)

    # -------------------------------------------------------------------------
    # STEP 4 - MASTER SKU LOOKUP (STYLE NAME + PAYOUT)
    # -------------------------------------------------------------------------
    logger.info("\n--- STEP 4: Master SKU Lookup (Style Name + Payout) ---")

    # Clean Master SKU keys (trim whitespace, lower-case for matching)
    master_df["_clean_master_sku"] = (
        master_df[master_sku_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lstrip("'")
        .str.strip()
        .str.lower()
    )
    master_df["_clean_style_name"] = master_df[master_style_col].fillna("").astype(str).str.strip()
    master_df["_clean_payout"] = clean_amount_series(master_df[master_payout_col], "Master SKU Payout", logger)

    # Check for duplicate SKU Name in Master: keep first and log warning
    master_dup_mask = master_df.duplicated(subset=["_clean_master_sku"], keep=False) & (master_df["_clean_master_sku"] != "")
    master_dup_count = int(master_dup_mask.sum())
    if master_dup_count > 0:
        unique_master_dups = master_df.loc[master_dup_mask, "_clean_master_sku"].nunique()
        logger.warning(
            f"Master SKU file contains {master_dup_count} duplicate rows ({unique_master_dups} unique SKU keys). "
            f"Keeping the first occurrence and ignoring subsequent duplicates."
        )
        sample_master_dups = master_df.loc[master_dup_mask, "_clean_master_sku"].unique()[:5]
        logger.warning(f"  -> Sample duplicate Master SKU keys: {list(sample_master_dups)}")

    master_dedup = master_df.drop_duplicates(subset=["_clean_master_sku"], keep="first")
    master_lookup = master_dedup[["_clean_master_sku", "_clean_style_name", "_clean_payout"]]

    # Base table normalized SKU for matching
    base_with_settlement["_match_sku"] = (
        base_with_settlement["seller sku code"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    base_with_master = base_with_settlement.merge(
        master_lookup,
        left_on="_match_sku",
        right_on="_clean_master_sku",
        how="left"
    )
    base_with_master.sort_values(by="_row_order", inplace=True)

    # No match (including when seller sku code is "NOT FOUND") -> Style Name = "NOT FOUND", Payout = 0.0
    base_with_master["Style Name"] = base_with_master["_clean_style_name"].fillna("NOT FOUND")
    base_with_master.loc[base_with_master["Style Name"].str.strip() == "", "Style Name"] = "NOT FOUND"
    base_with_master.loc[base_with_master["seller sku code"] == "NOT FOUND", "Style Name"] = "NOT FOUND"

    base_with_master["Payout"] = base_with_master["_clean_payout"].fillna(0.0)
    base_with_master.loc[base_with_master["seller sku code"] == "NOT FOUND", "Payout"] = 0.0

    # Collect de-duplicated list of seller sku codes not found in master
    missing_from_master_mask = (base_with_master["seller sku code"] != "NOT FOUND") & (base_with_master["Style Name"] == "NOT FOUND")
    missing_skus_list = sorted(list(set(base_with_master.loc[missing_from_master_mask, "seller sku code"])))

    matched_master_count = int(((base_with_master["seller sku code"] != "NOT FOUND") & (base_with_master["Style Name"] != "NOT FOUND")).sum())
    total_payout_sum = float(base_with_master["Payout"].sum())

    # Construct EXACT Final Sheet columns in exact required order matching sample reference:
    # order release id | seller sku code | Style Name | return_type | QTY | RETURN QTY | SALE QTY  | Payout | NET PRODUCT COST  | Settled_Amount | PROFIT
    qty_series = pd.Series(1, index=base_with_master.index, dtype=int)
    return_qty_series = base_with_master["mapped_return_type"].apply(
        lambda r: 0 if str(r).strip().lower() == "delivered" else 1
    )
    sale_qty_series = qty_series - return_qty_series
    payout_series = base_with_master["Payout"]
    net_cost_series = payout_series * sale_qty_series
    settled_series = base_with_master["Settled_Amount"]
    profit_series = settled_series - net_cost_series

    final_sheet_df = pd.DataFrame({
        "order release id": base_with_master[flow_sale_code_col],
        "seller sku code": base_with_master["seller sku code"],
        "Style Name": base_with_master["Style Name"],
        "return_type": base_with_master["mapped_return_type"],
        "QTY": qty_series,
        "RETURN QTY": return_qty_series,
        "SALE QTY ": sale_qty_series,
        "Payout": payout_series,
        "NET PRODUCT COST ": net_cost_series,
        "Settled_Amount": settled_series,
        "PROFIT": profit_series
    })

    sum_final_settled = float(final_sheet_df["Settled_Amount"].sum())
    total_net_cost_sum = float(final_sheet_df["NET PRODUCT COST "].sum())
    total_profit_sum = float(final_sheet_df["PROFIT"].sum())
    total_sale_qty = int(final_sheet_df["SALE QTY "].sum())
    total_return_qty = int(final_sheet_df["RETURN QTY"].sum())

    # -------------------------------------------------------------------------
    # VALIDATION & RECONCILIATION AUDIT
    # -------------------------------------------------------------------------
    logger.info("-" * 80)
    logger.info("VALIDATION & AUDIT SUMMARY:")
    logger.info(f"  * Seller Orders Report Total Rows : {orders_count:,}")
    logger.info(f"  * Order Flow Report Total Rows    : {flow_count:,}")
    logger.info(f"  * Prepaid Settlement Rows         : {prepaid_count:,}")
    logger.info(f"  * Postpaid Settlement Rows        : {postpaid_count:,}")
    logger.info(f"  * Master SKU Report Total Rows    : {master_count:,}")
    logger.info(f"  * Seller SKU Matched Rows (Orders): {matched_sku_count:,} ({matched_sku_count / flow_count * 100:.2f}%)")
    logger.info(f"  * Seller SKU NOT FOUND Rows       : {not_found_sku_count:,}")
    logger.info(f"  * Master SKU Matched Rows         : {matched_master_count:,}")
    logger.info(f"  * SKUs Missing from Master Count  : {len(missing_skus_list):,} unique SKUs")
    if missing_skus_list:
        logger.warning(f"  * List of SKUs missing from Master file ({len(missing_skus_list)} total):")
        for s in missing_skus_list[:20]:
            logger.warning(f"      - '{s}'")
        if len(missing_skus_list) > 20:
            logger.warning(f"      ... and {len(missing_skus_list) - 20} more missing SKUs.")
    else:
        logger.info("  * All found Seller SKUs matched cleanly in Master file!")

    logger.info(f"  * Seller Key Duplicates Handled   : {dup_count:,}")
    logger.info(f"  * Repeated sale_order_code Count  : {repeated_count:,} values ({total_repeated_rows:,} rows)")
    logger.info(f"  * Unmapped return_type Count      : {unmapped_count:,}")
    logger.info("  * Final Sheet Rows per return_type:")
    for status, cnt in final_sheet_df["return_type"].value_counts().items():
        logger.info(f"      - {status}: {cnt:,} ({cnt / flow_count * 100:.2f}%)")

    logger.info("\nSETTLEMENT & FINANCIAL TOTALS:")
    logger.info(f"  * Total Prepaid Settlement Amount : Rs. {total_prepaid_amount:,.2f}")
    logger.info(f"  * Total Postpaid Settlement Amount: Rs. {total_postpaid_amount:,.2f}")
    logger.info(f"  * Combined Settlement Total       : Rs. {total_combined_settlement:,.2f}")
    logger.info(f"  * Final Sheet Settled Amount Sum  : Rs. {sum_final_settled:,.2f}")
    logger.info(f"  * Final Sheet Total Payout Sum    : Rs. {total_payout_sum:,.2f}")

    # Reconciliation Check: sum of Settled_Amount in "Final" vs combined total
    settled_flow_release_ids = set(base_merged["_clean_sale_order_code"])
    all_settled_release_ids = set(settlement_stacked_df["order release id"])
    unlinked_settled_ids = all_settled_release_ids - settled_flow_release_ids
    unlinked_settlement_sum = float(settlement_stacked_df.loc[settlement_stacked_df["order release id"].isin(unlinked_settled_ids), "Settled_Amount"].sum())

    reconciliation_diff = round(total_combined_settlement - sum_final_settled, 2)
    if reconciliation_diff > 0.001:
        logger.warning(
            f"RECONCILIATION DISCREPANCY: Combined Settlement Total exceeds Final Report sum by Rs. {reconciliation_diff:,.2f}!"
        )
        logger.warning(
            f"Found {len(unlinked_settled_ids)} order release id(s) settled in Prepaid/Postpaid that did NOT match any order in Order Flow:"
        )
        for uid in sorted(list(unlinked_settled_ids))[:15]:
            amt = float(settlement_stacked_df.loc[settlement_stacked_df["order release id"] == uid, "Settled_Amount"].sum())
            logger.warning(f"  -> Unlinked Release ID: '{uid}' | Settled Amount: Rs. {amt:,.2f}")
        if len(unlinked_settled_ids) > 15:
            logger.warning(f"  ... and {len(unlinked_settled_ids) - 15} more unlinked release ID(s).")
        logger.warning(f"Total unlinked settlement funds: Rs. {unlinked_settlement_sum:,.2f}")
    else:
        logger.info("RECONCILIATION PASSED: 100% of settlement funds matched into Final report cleanly.")

    logger.info("-" * 80)

    # -------------------------------------------------------------------------
    # OUTPUT WORKBOOK GENERATION
    # -------------------------------------------------------------------------
    os.makedirs(output_dir, exist_ok=True)
    today_str = datetime.now().strftime("%Y-%m-%d")
    final_xlsx_path = os.path.join(output_dir, f"Final_Report_{today_str}.xlsx")

    actual_xlsx_path = write_excel_final_report(
        xlsx_path=final_xlsx_path,
        final_df=final_sheet_df,
        settlement_data_df=settlement_stacked_df,
        settlement_pivot_df=pivot_export_df,
        logger=logger,
        include_settlement_sheets=include_settlement_sheets
    )
    final_xlsx_path = actual_xlsx_path or final_xlsx_path
    logger.info(f"Successfully generated Final Excel Report: {final_xlsx_path}")
    logger.info("=" * 80)
    logger.info("PIPELINE EXECUTION COMPLETED SUCCESSFULLY.")
    logger.info("=" * 80)

    return {
        "final_xlsx_path": final_xlsx_path,
        "final_sheet_df": final_sheet_df,
        "settlement_stacked_df": settlement_stacked_df,
        "pivot_export_df": pivot_export_df,
        "orders_count": orders_count,
        "flow_count": flow_count,
        "prepaid_count": prepaid_count,
        "postpaid_count": postpaid_count,
        "master_count": master_count,
        "matched_sku_count": matched_sku_count,
        "not_found_sku_count": not_found_sku_count,
        "matched_master_count": matched_master_count,
        "missing_skus_list": missing_skus_list,
        "total_prepaid_amount": total_prepaid_amount,
        "total_postpaid_amount": total_postpaid_amount,
        "total_combined_settlement": total_combined_settlement,
        "sum_final_settled": sum_final_settled,
        "total_payout_sum": total_payout_sum,
        "total_net_cost_sum": total_net_cost_sum,
        "total_profit_sum": total_profit_sum,
        "total_sale_qty": total_sale_qty,
        "total_return_qty": total_return_qty,
        "unlinked_settled_count": len(unlinked_settled_ids),
        "unlinked_settlement_sum": unlinked_settlement_sum,
        "repeated_codes_count": repeated_count,
        "log_entries": getattr(logger, "entries", [])
    }


def main():
    parser = argparse.ArgumentParser(description="Order Flow & Settlement Reconciliation Pipeline (5-File Suite)")
    parser.add_argument("--orders-file", type=str, default=None, help="Explicit path to Seller Orders CSV/Excel")
    parser.add_argument("--flow-file", type=str, default=None, help="Explicit path to Order Flow CSV/Excel")
    parser.add_argument("--prepaid-file", type=str, default=None, help="Explicit path to Prepaid Settlement CSV/Excel")
    parser.add_argument("--postpaid-file", type=str, default=None, help="Explicit path to Postpaid Settlement CSV/Excel")
    parser.add_argument("--master-file", type=str, default=None, help="Explicit path to Master SKU CSV/Excel")
    parser.add_argument("--input-dir", type=str, default=DEFAULT_INPUT_DIR, help="Input directory")
    parser.add_argument("--output-dir", type=str, default=DEFAULT_OUTPUT_DIR, help="Output directory")
    parser.add_argument("--include-settlement-sheets", action="store_true", help="Include Settlement_Data and Settlement_Pivot sheets")
    args = parser.parse_args()

    process_order_flow_pipeline(
        orders_path=args.orders_file,
        flow_path=args.flow_file,
        prepaid_path=args.prepaid_file,
        postpaid_path=args.postpaid_file,
        master_path=args.master_file,
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        include_settlement_sheets=args.include_settlement_sheets
    )


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
FastAPI Web Application Server for Order Flow Lookup & Status Enrichment Suite (5-File Architecture).
Human-centric enterprise backend with robust error reporting, 5-file management,
and live Excel report processing.
"""
import os
import sys
import glob
import re
import io
import shutil
from datetime import datetime
from typing import Optional, Dict, Any, List

from fastapi import FastAPI, File, UploadFile, Query, HTTPException, Body, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd

# Import pipeline components from run.py
from run import (
    process_order_flow_pipeline,
    find_input_files,
    load_file_as_text,
    DualLogger,
    RETURN_TYPE_STATUS_MAPPING,
    DEFAULT_BLANK_STATUS,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# On Vercel / serverless runtimes, local filesystem is read-only except /tmp
IS_VERCEL = bool(os.environ.get("VERCEL"))
if IS_VERCEL:
    INPUT_DIR = os.path.join("/tmp", "input")
    OUTPUT_DIR = os.path.join("/tmp", "output")
    LOGS_DIR = os.path.join("/tmp", "logs")
else:
    INPUT_DIR = os.path.join(BASE_DIR, "input")
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")
    LOGS_DIR = os.path.join(BASE_DIR, "logs")

try:
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(LOGS_DIR, exist_ok=True)
except Exception:
    pass

app = FastAPI(
    title="OrderFlow Studio",
    description="Enterprise Order Flow & 5-File Settlement Reconciliation Suite",
    version="3.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.exception_handler(Exception)
async def unhandled_exception_handler(request, exc):
    return JSONResponse(
        status_code=500,
        content={"success": False, "error": f"Internal Server Error: {str(exc)}"}
    )

STATE: Dict[str, Any] = {
    "latest_result": None,
    "last_run_timestamp": None,
    "active_orders_file": None,
    "active_flow_file": None,
    "active_prepaid_file": None,
    "active_postpaid_file": None,
    "active_master_file": None,
    "custom_status_mapping": dict(RETURN_TYPE_STATUS_MAPPING),
    "blank_status": DEFAULT_BLANK_STATUS,
}


def format_bytes(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


def inspect_file(filepath: str) -> Dict[str, Any]:
    """Inspects a CSV or Excel file to determine its columns, row count, and role."""
    fname = os.path.basename(filepath)
    fsize = os.path.getsize(filepath)
    mtime = datetime.fromtimestamp(os.path.getmtime(filepath)).strftime("%Y-%m-%d %H:%M:%S")

    cols = []
    original_cols = []
    row_count = 0

    ext = os.path.splitext(fname)[1].lower()

    if ext in (".xlsx", ".xls"):
        try:
            sample_df = pd.read_excel(filepath, nrows=5, dtype=str)
            original_cols = [str(c).strip() for c in sample_df.columns]
            cols = [c.lower() for c in original_cols]
            # Fast row estimation for Excel: open with openpyxl
            import openpyxl
            wb = openpyxl.load_workbook(filepath, read_only=True)
            ws = wb.active
            row_count = max(0, (ws.max_row or 0) - 1)
            wb.close()
        except Exception:
            row_count = 0
    else:
        encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
        for enc in encodings:
            try:
                sample_df = pd.read_csv(filepath, nrows=5, encoding=enc, dtype=str)
                original_cols = [str(c).strip() for c in sample_df.columns]
                cols = [c.lower() for c in original_cols]
                break
            except Exception:
                continue

        # Line count for CSV
        try:
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                row_count = max(0, sum(1 for _ in f) - 1)
        except Exception:
            row_count = 0

    fname_lower = fname.lower()
    is_master = "master" in fname_lower or ("payout" in cols and ("style name" in cols or "style code" in cols))
    is_orders = ("order release id" in cols and "seller sku code" in cols) or ("seller_orders_report" in fname_lower) or ("seller" in fname_lower and "orders" in fname_lower)
    is_flow = ("sale_order_code" in cols and "return_type" in cols) or (fname_lower.startswith("part-")) or ("flow" in fname_lower)

    # Settlement detection
    has_settle_cols = ("order release id" in cols or "order line id" in cols) and ("settled_amount" in cols or "settled amount" in cols)
    is_prepaid = ("prepaid" in fname_lower or "pre-paid" in fname_lower) and ("postpaid" not in fname_lower and "cod" not in fname_lower)
    is_postpaid = "postpaid" in fname_lower or "post-paid" in fname_lower or "cod" in fname_lower

    if is_master:
        role = "master_sku"
    elif is_prepaid or (has_settle_cols and is_prepaid):
        role = "prepaid_settlement"
    elif is_postpaid or (has_settle_cols and is_postpaid):
        role = "postpaid_settlement"
    elif has_settle_cols:
        role = "generic_settlement"
    elif is_orders and not is_flow:
        role = "seller_orders"
    elif is_flow and not is_orders:
        role = "order_flow"
    else:
        role = "unknown"

    return {
        "filename": fname,
        "filepath": filepath,
        "size_bytes": fsize,
        "size_formatted": format_bytes(fsize),
        "modified": mtime,
        "row_count": row_count,
        "role": role,
        "columns": original_cols,
    }


def scan_input_directory() -> Dict[str, Any]:
    """Scans input/ directory and categorizes all available files across the 5 required types."""
    raw_paths = glob.glob(os.path.join(INPUT_DIR, "*"))
    file_paths = [f for f in raw_paths if f.lower().endswith((".csv", ".xlsx", ".xls"))]

    orders_files = []
    flow_files = []
    prepaid_files = []
    postpaid_files = []
    master_files = []
    all_files = []

    for path in file_paths:
        info = inspect_file(path)
        all_files.append(info)
        fname_lower = info["filename"].lower()
        role = info["role"]

        if role == "master_sku" or "master" in fname_lower:
            master_files.append(info)
        elif role == "seller_orders" or "seller_orders_report" in fname_lower or ("seller" in fname_lower and "orders" in fname_lower):
            orders_files.append(info)
        elif role == "order_flow" or fname_lower.startswith("part-") or "flow" in fname_lower:
            flow_files.append(info)
        elif role == "prepaid_settlement" or ("prepaid" in fname_lower and "postpaid" not in fname_lower):
            prepaid_files.append(info)
        elif role == "postpaid_settlement" or "postpaid" in fname_lower or "cod" in fname_lower:
            postpaid_files.append(info)
        elif role == "generic_settlement":
            if not prepaid_files:
                prepaid_files.append(info)
            elif not postpaid_files:
                postpaid_files.append(info)

    # Sort files by newest modification time first
    orders_files.sort(key=lambda x: x["modified"], reverse=True)
    flow_files.sort(key=lambda x: x["modified"], reverse=True)
    prepaid_files.sort(key=lambda x: x["modified"], reverse=True)
    postpaid_files.sort(key=lambda x: x["modified"], reverse=True)
    master_files.sort(key=lambda x: x["modified"], reverse=True)
    all_files.sort(key=lambda x: x["modified"], reverse=True)

    detected_orders = orders_files[0] if orders_files else None
    detected_flow = flow_files[0] if flow_files else None
    detected_prepaid = prepaid_files[0] if prepaid_files else None
    detected_postpaid = postpaid_files[0] if postpaid_files else None
    detected_master = master_files[0] if master_files else None

    ready = bool(detected_orders and detected_flow and detected_prepaid and detected_postpaid and detected_master)

    return {
        "orders_files": orders_files,
        "flow_files": flow_files,
        "prepaid_files": prepaid_files,
        "postpaid_files": postpaid_files,
        "master_files": master_files,
        "all_files": all_files,
        "detected_orders": detected_orders,
        "detected_flow": detected_flow,
        "detected_prepaid": detected_prepaid,
        "detected_postpaid": detected_postpaid,
        "detected_master": detected_master,
        "ready": ready
    }


@app.get("/api/status")
def get_system_status():
    """Returns current system readiness, detected files, and latest output metadata."""
    scan = scan_input_directory()

    excel_outputs = sorted(glob.glob(os.path.join(OUTPUT_DIR, "Final_Report_*.xlsx")), reverse=True)
    if not excel_outputs:
        excel_outputs = sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), reverse=True)
    csv_outputs = sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.csv")), reverse=True)

    latest_excel = os.path.basename(excel_outputs[0]) if excel_outputs else None
    latest_csv = os.path.basename(csv_outputs[0]) if csv_outputs else None

    return {
        **scan,
        "latest_excel": latest_excel,
        "latest_csv": latest_csv,
        "has_cached_result": STATE["latest_result"] is not None,
        "last_run_timestamp": STATE["last_run_timestamp"],
        "default_mappings": STATE["custom_status_mapping"],
        "default_blank_status": STATE["blank_status"],
    }


@app.post("/api/select-file")
def select_file(payload: Dict[str, str] = Body(...)):
    """Allows selecting a specific file for any role from the input folder."""
    role = payload.get("role")
    filename = payload.get("filename")

    if not filename or not os.path.exists(os.path.join(INPUT_DIR, filename)):
        raise HTTPException(status_code=400, detail=f"File '{filename}' does not exist in input directory.")

    if role == "seller_orders":
        STATE["active_orders_file"] = filename
    elif role == "order_flow":
        STATE["active_flow_file"] = filename
    elif role == "prepaid_settlement":
        STATE["active_prepaid_file"] = filename
    elif role == "postpaid_settlement":
        STATE["active_postpaid_file"] = filename
    elif role == "master_sku":
        STATE["active_master_file"] = filename
    else:
        raise HTTPException(status_code=400, detail="Invalid role specified.")

    return get_system_status()


@app.post("/api/upload")
async def upload_file(
    files: List[UploadFile] = File(...),
    target_role: Optional[str] = Form(None)
):
    """
    Upload one or multiple CSV or Excel files.
    Supports auto-detection and role-targeted uploading for all 5 types.
    """
    results = []

    for file in files:
        fname = os.path.basename(file.filename)
        fname_lower = fname.lower()
        if not (fname_lower.endswith(".csv") or fname_lower.endswith(".xlsx") or fname_lower.endswith(".xls")):
            results.append({
                "filename": file.filename,
                "success": False,
                "error": "Only CSV (.csv) and Excel (.xlsx, .xls) files are supported."
            })
            continue

        raw_contents = await file.read()
        
        # Save file directly with appropriate name
        clean_filename = fname
        c_lower = clean_filename.lower()

        if target_role == "master_sku" and "master" not in c_lower:
            clean_filename = f"Master_{clean_filename}"
        elif target_role == "prepaid_settlement" and "prepaid" not in c_lower:
            clean_filename = f"Prepaid_{clean_filename}"
        elif target_role == "postpaid_settlement" and "postpaid" not in c_lower:
            clean_filename = f"Postpaid_{clean_filename}"
        elif target_role == "seller_orders" and "seller" not in c_lower:
            clean_filename = f"Seller_Orders_{clean_filename}"
        elif target_role == "order_flow" and not c_lower.startswith("part-") and "flow" not in c_lower:
            clean_filename = f"part-00000_{clean_filename}"

        saved_path = os.path.join(INPUT_DIR, clean_filename)
        with open(saved_path, "wb") as f:
            f.write(raw_contents)

        info = inspect_file(saved_path)

        # Assign role
        valid_roles = ("seller_orders", "order_flow", "prepaid_settlement", "postpaid_settlement", "master_sku")
        assigned_role = target_role if target_role in valid_roles else info["role"]

        if assigned_role == "generic_settlement":
            scan_now = scan_input_directory()
            if not scan_now.get("detected_prepaid"):
                assigned_role = "prepaid_settlement"
                if "prepaid" not in clean_filename.lower():
                    new_name = f"Prepaid_{clean_filename}"
                    new_path = os.path.join(INPUT_DIR, new_name)
                    os.rename(saved_path, new_path)
                    saved_path = new_path
                    clean_filename = new_name
            else:
                assigned_role = "postpaid_settlement"
                if "postpaid" not in clean_filename.lower():
                    new_name = f"Postpaid_{clean_filename}"
                    new_path = os.path.join(INPUT_DIR, new_name)
                    os.rename(saved_path, new_path)
                    saved_path = new_path
                    clean_filename = new_name

        if assigned_role == "seller_orders":
            STATE["active_orders_file"] = clean_filename
        elif assigned_role == "order_flow":
            STATE["active_flow_file"] = clean_filename
        elif assigned_role == "prepaid_settlement":
            STATE["active_prepaid_file"] = clean_filename
        elif assigned_role == "postpaid_settlement":
            STATE["active_postpaid_file"] = clean_filename
        elif assigned_role == "master_sku":
            STATE["active_master_file"] = clean_filename

        results.append({
            **info,
            "filename": clean_filename,
            "success": True,
            "assigned_role": assigned_role
        })

    scan = scan_input_directory()
    return {
        "uploaded": results,
        "status": scan
    }


@app.post("/api/remove-file")
def remove_file(payload: Dict[str, Any] = Body(...)):
    """
    Cancels and removes an uploaded file by role ('seller_orders', 'order_flow',
    'prepaid_settlement', 'postpaid_settlement', 'master_sku') or by filename.
    Moves the file out of input/ into output/archive/ so it is no longer detected.
    """
    role = payload.get("role")
    filename = payload.get("filename")

    scan = scan_input_directory()
    target_filename = None

    if role == "seller_orders":
        if scan.get("detected_orders"):
            target_filename = scan["detected_orders"]["filename"]
        STATE["active_orders_file"] = None
    elif role == "order_flow":
        if scan.get("detected_flow"):
            target_filename = scan["detected_flow"]["filename"]
        STATE["active_flow_file"] = None
    elif role == "prepaid_settlement":
        if scan.get("detected_prepaid"):
            target_filename = scan["detected_prepaid"]["filename"]
        STATE["active_prepaid_file"] = None
    elif role == "postpaid_settlement":
        if scan.get("detected_postpaid"):
            target_filename = scan["detected_postpaid"]["filename"]
        STATE["active_postpaid_file"] = None
    elif role == "master_sku":
        if scan.get("detected_master"):
            target_filename = scan["detected_master"]["filename"]
        STATE["active_master_file"] = None
    elif filename:
        target_filename = filename

    if target_filename:
        target_path = os.path.join(INPUT_DIR, target_filename)
        if os.path.exists(target_path):
            archive_dir = os.path.join(OUTPUT_DIR, "archive")
            os.makedirs(archive_dir, exist_ok=True)
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest_path = os.path.join(archive_dir, f"removed_{timestamp_str}_{target_filename}")
            try:
                shutil.move(target_path, dest_path)
            except Exception:
                try:
                    os.remove(target_path)
                except Exception:
                    pass

    # Clear cached result when input files change
    STATE["latest_result"] = None

    return get_system_status()


@app.post("/api/run")
def run_pipeline(payload: Dict[str, Any] = Body(default={})):
    """Executes the reconciliation pipeline using auto-detected or specified files."""
    log_path = os.path.join(LOGS_DIR, "log.txt")
    logger = DualLogger(log_path)

    try:
        scan = scan_input_directory()
        orders_filename = payload.get("orders_filename") or (scan["detected_orders"]["filename"] if scan["detected_orders"] else None)
        flow_filename = payload.get("flow_filename") or (scan["detected_flow"]["filename"] if scan["detected_flow"] else None)
        prepaid_filename = payload.get("prepaid_filename") or (scan["detected_prepaid"]["filename"] if scan["detected_prepaid"] else None)
        postpaid_filename = payload.get("postpaid_filename") or (scan["detected_postpaid"]["filename"] if scan["detected_postpaid"] else None)
        master_filename = payload.get("master_filename") or (scan["detected_master"]["filename"] if scan["detected_master"] else None)

        missing = []
        if not orders_filename or not os.path.exists(os.path.join(INPUT_DIR, orders_filename)):
            missing.append("1. Seller Orders Report")
        if not flow_filename or not os.path.exists(os.path.join(INPUT_DIR, flow_filename)):
            missing.append("2. Order Flow Report")
        if not prepaid_filename or not os.path.exists(os.path.join(INPUT_DIR, prepaid_filename)):
            missing.append("3. Prepaid Settlement")
        if not postpaid_filename or not os.path.exists(os.path.join(INPUT_DIR, postpaid_filename)):
            missing.append("4. Postpaid Settlement")
        if not master_filename or not os.path.exists(os.path.join(INPUT_DIR, master_filename)):
            missing.append("5. Master SKU Report")

        if missing:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": f"Cannot execute reconciliation: Missing file(s): {', '.join(missing)}. Please upload all 5 files."
                }
            )

        orders_path = os.path.join(INPUT_DIR, orders_filename)
        flow_path = os.path.join(INPUT_DIR, flow_filename)
        prepaid_path = os.path.join(INPUT_DIR, prepaid_filename)
        postpaid_path = os.path.join(INPUT_DIR, postpaid_filename)
        master_path = os.path.join(INPUT_DIR, master_filename)

        custom_mapping = payload.get("status_mapping") or STATE["custom_status_mapping"]
        blank_status = payload.get("blank_status") or STATE["blank_status"]

        result = process_order_flow_pipeline(
            orders_path=orders_path,
            flow_path=flow_path,
            prepaid_path=prepaid_path,
            postpaid_path=postpaid_path,
            master_path=master_path,
            input_dir=INPUT_DIR,
            output_dir=OUTPUT_DIR,
            log_path=log_path,
            mapping_dict=custom_mapping,
            blank_status=blank_status
        )

        STATE["latest_result"] = result
        STATE["last_run_timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Convert final_sheet_df to preview records
        final_sheet_df = result["final_sheet_df"]
        preview_records = final_sheet_df.head(100).to_dict(orient="records")

        return {
            "success": True,
            "timestamp": STATE["last_run_timestamp"],
            "orders_rows": result["orders_count"],
            "flow_rows": result["flow_count"],
            "prepaid_rows": result["prepaid_count"],
            "postpaid_rows": result["postpaid_count"],
            "master_rows": result["master_count"],
            "matched_count": result["matched_sku_count"],
            "not_found_count": result["not_found_sku_count"],
            "matched_master_count": result["matched_master_count"],
            "missing_skus_count": len(result["missing_skus_list"]),
            "match_rate": round(result["matched_sku_count"] / max(result["flow_count"], 1) * 100, 2),
            "total_prepaid_amount": result.get("total_prepaid_amount", 0.0),
            "total_postpaid_amount": result.get("total_postpaid_amount", 0.0),
            "total_combined_settlement": result["total_combined_settlement"],
            "sum_final_settled": result["sum_final_settled"],
            "total_payout_sum": result["total_payout_sum"],
            "total_net_cost_sum": result.get("total_net_cost_sum", 0.0),
            "total_profit_sum": result.get("total_profit_sum", 0.0),
            "unlinked_settled_count": result["unlinked_settled_count"],
            "unlinked_settlement_sum": result["unlinked_settlement_sum"],
            "log_entries": result["log_entries"],
            "preview_rows": preview_records,
            "total_records": len(final_sheet_df),
            "xlsx_download_url": "/api/download/excel",
            "log_download_url": "/api/download/log"
        }

    except KeyError as e:
        err_msg = f"Column Validation Error: {str(e)}"
        logger.error(err_msg)
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": err_msg, "log_entries": getattr(logger, "entries", [])}
        )
    except Exception as e:
        err_msg = f"Processing Error: {str(e)}"
        logger.error(err_msg)
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": err_msg, "log_entries": getattr(logger, "entries", [])}
        )


@app.get("/api/data")
def query_enriched_data(
    page: int = Query(1, ge=1),
    limit: int = Query(25, ge=5, le=200),
    search: Optional[str] = Query(None),
    return_type: Optional[str] = Query(None),
    match_status: Optional[str] = Query("all")
):
    """High performance server-side search, filtering, and pagination for final report records."""
    if not STATE["latest_result"]:
        raise HTTPException(status_code=404, detail="No enrichment pipeline results available yet. Run pipeline first.")

    res = STATE["latest_result"]
    if "final_sheet_df" in res and hasattr(res["final_sheet_df"], "to_dict"):
        records = res["final_sheet_df"].to_dict(orient="records")
    else:
        records = res.get("preview_rows", [])

    filtered = records

    if match_status == "matched":
        filtered = [r for r in filtered if r.get("seller sku code") != "NOT FOUND"]
    elif match_status == "not_found":
        filtered = [r for r in filtered if r.get("seller sku code") == "NOT FOUND"]

    if return_type and return_type != "all":
        filtered = [r for r in filtered if str(r.get("return_type", "")).lower() == return_type.lower()]

    if search:
        s = search.strip().lower()
        filtered = [
            r for r in filtered
            if any(s in str(v).lower() for v in r.values())
        ]

    total_count = len(filtered)
    total_pages = max(1, (total_count + limit - 1) // limit)
    start_idx = (page - 1) * limit
    end_idx = start_idx + limit
    page_data = filtered[start_idx:end_idx]

    return {
        "page": page,
        "limit": limit,
        "total_records": total_count,
        "total_pages": total_pages,
        "data": page_data
    }


@app.get("/api/download/{file_type}")
def download_output_file(file_type: str):
    """Download generated files: excel or log."""
    if file_type == "excel":
        latest_path = STATE["latest_result"].get("final_xlsx_path") if STATE["latest_result"] else None
        if latest_path and os.path.exists(latest_path):
            path = latest_path
        else:
            excel_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "Final_Report_*.xlsx")), reverse=True)
            if not excel_files:
                excel_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), reverse=True)
            if not excel_files:
                raise HTTPException(status_code=404, detail="No Excel report found.")
            path = excel_files[0]
        return FileResponse(
            path,
            filename=os.path.basename(path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    elif file_type == "log":
        log_path = os.path.join(LOGS_DIR, "log.txt")
        if not os.path.exists(log_path):
            raise HTTPException(status_code=404, detail="No audit log found.")
        return FileResponse(
            log_path,
            filename="audit_log.txt",
            media_type="text/plain"
        )

    raise HTTPException(status_code=400, detail="Invalid file type. Choose excel or log.")


# Mount static files
STATIC_DIR = os.path.join(BASE_DIR, "static")
if os.path.exists(STATIC_DIR):
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    port = 8000
    print(f"\n[SERVER] Starting OrderFlow Intelligence Web App on http://localhost:{port}\n")
    uvicorn.run("web_server:app", host="127.0.0.1", port=port, reload=False)

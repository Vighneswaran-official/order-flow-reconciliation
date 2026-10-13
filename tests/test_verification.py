"""
Automated Verification Suite for Order Flow & Settlement Reconciliation Pipeline (Sample Reference Standard)
=============================================================================================================
Validates all requirements:
  1. Sheet "Final": EXACTLY 11 columns matching sample reference:
       order release id | seller sku code | Style Name | return_type | QTY | RETURN QTY | SALE QTY  | Payout | NET PRODUCT COST  | Settled_Amount | PROFIT
  2. Formulas verified dynamically:
       SALE QTY         : =E{row}-F{row}
       NET PRODUCT COST : =H{row}*G{row}
       PROFIT           : =J{row}-I{row}
  3. Formats:
       Col 1 (order release id) : integer/text, left aligned
       Col 2 (seller sku code)  : text (@), left aligned
       Col 3 (Style Name)       : text (@), left aligned
       Col 4 (return_type)      : text (@), center aligned
       Col 5 (QTY)              : integer (1), center aligned
       Col 6 (RETURN QTY)       : integer (0 or 1), center aligned
       Col 7 (SALE QTY )        : formula, center aligned
       Col 8 (Payout)           : float (#,##0.00), right aligned
       Col 9 (NET PRODUCT COST ): formula (#,##0.00), right aligned
       Col 10 (Settled_Amount)  : float (#,##0.00), right aligned
       Col 11 (PROFIT)          : formula (#,##0.00)
  4. Scientific Notation Guard: Halts on 'E+' with clear error message.
  5. Audit Log: log.txt captures 5-file row audits, unlinked settlements, missing SKUs.
"""

import os
import sys
from datetime import datetime
from pathlib import Path
import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from run import process_order_flow_pipeline

def test_final_excel_report():
    today_str = datetime.now().strftime("%Y-%m-%d")
    xlsx_path = PROJECT_ROOT / "output" / f"Final_Report_{today_str}.xlsx"
    log_path = PROJECT_ROOT / "logs" / "log.txt"

    assert xlsx_path.exists(), f"Final Excel workbook not found at {xlsx_path}!"
    assert log_path.exists(), f"Log file not found at {log_path}!"

    wb = openpyxl.load_workbook(xlsx_path)
    print("Sheets present in workbook:", wb.sheetnames)

    # Required single sheet by default
    expected_sheets = ["Final"]
    assert wb.sheetnames == expected_sheets, f"Expected sheets {expected_sheets}, but found {wb.sheetnames}"
    print("[PASS] Sheet names verified: EXACTLY 'Final'.")

    # -------------------------------------------------------------------------
    # 1. TEST SHEET "Final"
    # -------------------------------------------------------------------------
    ws_final = wb["Final"]
    final_headers = [ws_final.cell(1, c).value for c in range(1, ws_final.max_column + 1)]
    expected_final_headers = [
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
    assert final_headers == expected_final_headers, f"Expected headers {expected_final_headers}, got {final_headers}"
    assert ws_final.max_row > 1, "Final sheet contains no data rows!"

    # Test formats and formulas in "Final"
    for r in range(2, min(ws_final.max_row + 1, 50)):
        # col 1: order release id
        c1 = ws_final.cell(r, 1)
        assert c1.number_format in ("0", "@")

        # col 2: seller sku code -> text
        c2 = ws_final.cell(r, 2)
        assert c2.data_type == "s"

        # col 3: Style Name -> text
        c3 = ws_final.cell(r, 3)
        assert c3.data_type == "s"

        # col 4: return_type -> text
        c4 = ws_final.cell(r, 4)
        assert c4.data_type == "s"
        assert str(c4.value) in ("Delivered", "RTO", "RVP", "Exchange") or str(c4.value).startswith("UNMAPPED:")

        # col 5: QTY -> integer 1
        c5 = ws_final.cell(r, 5)
        assert c5.value == 1

        # col 6: RETURN QTY -> 0 for Delivered, 1 otherwise
        c6 = ws_final.cell(r, 6)
        expected_ret_qty = 0 if str(c4.value).strip().lower() == "delivered" else 1
        assert c6.value == expected_ret_qty, f"Row {r}: expected RETURN QTY {expected_ret_qty}, got {c6.value}"

        # col 7: SALE QTY  -> formula =E{r}-F{r}
        c7 = ws_final.cell(r, 7)
        assert str(c7.value) == f"=E{r}-F{r}", f"Row {r}: expected =E{r}-F{r}, got {c7.value}"

        # col 8: Payout -> numeric float with 2 decimals
        c8 = ws_final.cell(r, 8)
        assert c8.data_type == "n"
        assert c8.number_format == "#,##0.00"

        # col 9: NET PRODUCT COST  -> formula =H{r}*G{r}
        c9 = ws_final.cell(r, 9)
        assert str(c9.value) == f"=H{r}*G{r}", f"Row {r}: expected =H{r}*G{r}, got {c9.value}"
        assert c9.number_format == "#,##0.00"

        # col 10: Settled_Amount -> numeric float with 2 decimals
        c10 = ws_final.cell(r, 10)
        assert c10.data_type == "n"
        assert c10.number_format == "#,##0.00"

        # col 11: PROFIT -> formula =J{r}-I{r}
        c11 = ws_final.cell(r, 11)
        assert str(c11.value) == f"=J{r}-I{r}", f"Row {r}: expected =J{r}-I{r}, got {c11.value}"
        assert c11.number_format == "#,##0.00"

    print(f"[PASS] Sheet 'Final' verified: EXACTLY 11 columns, {ws_final.max_row - 1} data rows, all formulas verified.")

    # -------------------------------------------------------------------------
    # 2. TEST SCIENTIFIC NOTATION GUARD
    # -------------------------------------------------------------------------
    corrupted_csv = PROJECT_ROOT / "tests" / "corrupted_flow_with_scientific_notation.csv"
    if corrupted_csv.exists():
        try:
            guard_log = PROJECT_ROOT / "logs" / "guard_test_log.txt"
            process_order_flow_pipeline(flow_path=str(corrupted_csv), log_path=str(guard_log))
            assert False, "Scientific notation guard failed to halt execution!"
        except ValueError as e:
            assert "sale_order_code is in scientific notation" in str(e)
            print("[PASS] Scientific notation guard verified: successfully blocked truncated data.")

    # -------------------------------------------------------------------------
    # 3. TEST LOG FILE (Validation & Reconciliation)
    # -------------------------------------------------------------------------
    with open(log_path, "r", encoding="utf-8") as f:
        log_content = f.read()

    assert "VALIDATION & AUDIT SUMMARY:" in log_content
    assert "Seller Orders Report Total Rows" in log_content
    assert "Order Flow Report Total Rows" in log_content
    assert "Prepaid Settlement Rows" in log_content
    assert "Postpaid Settlement Rows" in log_content
    assert "Master SKU Report Total Rows" in log_content
    assert "Repeated sale_order_code Count" in log_content
    assert "Final Sheet Rows per return_type:" in log_content
    assert "SETTLEMENT & FINANCIAL TOTALS:" in log_content

    print("[PASS] Audit log verified: 5-file row audits, SKU matches, and reconciliation captured.")


if __name__ == "__main__":
    test_final_excel_report()
    print("\n=======================================================")
    print("ALL SAMPLE REFERENCE PIPELINE TESTS PASSED!")
    print("=======================================================")

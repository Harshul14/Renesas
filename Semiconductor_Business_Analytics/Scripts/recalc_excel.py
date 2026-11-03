"""Recalculate workbook using Microsoft Excel COM, check for formula errors,
and save final versions to root and Final_Deliverables."""
import sys
import os
import shutil
from pathlib import Path
import win32com.client as win32

PROJECT = Path(__file__).resolve().parent.parent
RAW_PATH = str(PROJECT / "Exports" / "_build_raw.xlsx")
FINAL_ROOT = str(PROJECT / "Semiconductor_Business_Analytics.xlsx")
FINAL_DELIV = str(PROJECT / "Final_Deliverables" / "Semiconductor_Business_Analytics.xlsx")

def recalc_and_audit():
    print(f"Opening Excel COM to recalculate {RAW_PATH}...")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(RAW_PATH))
        excel.CalculateFullRebuild()

        # Audit errors across all worksheets
        errors_found = []
        print("\n--- WORKSHEET AUDIT ---")
        for ws in wb.Worksheets:
            name = ws.Name
            used = ws.UsedRange
            vals = used.Value
            err_count = 0
            if vals is not None:
                # vals can be a tuple of tuples
                if not isinstance(vals, (list, tuple)):
                    vals = [[vals]]
                for r_idx, row in enumerate(vals):
                    if not isinstance(row, (list, tuple)):
                        row = [row]
                    for c_idx, val in enumerate(row):
                        if isinstance(val, int) and val in (-2146826281, -2146826246, -2146826259, -2146826252, -2146826265, -2146826273):
                            # Excel COM error codes like #REF!, #NAME?, #VALUE!, #DIV/0!, #N/A
                            err_cell = f"{name}!R{r_idx+1}C{c_idx+1}"
                            errors_found.append((err_cell, val))
                            err_count += 1
                        elif isinstance(val, str) and val.startswith("#") and val.endswith("!"):
                            err_cell = f"{name}!R{r_idx+1}C{c_idx+1}"
                            errors_found.append((err_cell, val))
                            err_count += 1

            status = f"ERRORS: {err_count}" if err_count > 0 else "CLEAN"
            print(f"Sheet '{name}': {status} (UsedRange: {used.Address})")

        # Check key validation metrics from Data_Validation sheet
        try:
            ws_val = wb.Worksheets("Data_Validation")
            print("\n--- DATA VALIDATION SUMMARY ---")
            val_summary = ws_val.Range("C46").Value  # e.g., summary cell
            print(f"Validation Summary: {val_summary}")
        except Exception as e:
            print(f"Could not read validation summary: {e}")

        # Check Dashboard key figures
        try:
            ws_dash = wb.Worksheets("Dashboard")
            print("\n--- DASHBOARD HIGHLIGHTS ---")
            print(f"Headline Recommendation: {ws_dash.Range('A4').Value}")
            print(f"FY25 Total Revenue: {ws_dash.Range('A7').Value}")
            print(f"CAGR FY23-25: {ws_dash.Range('C7').Value}")
            print(f"Highest Expected Growth: {ws_dash.Range('E7').Value}")
            print(f"Highest OM: {ws_dash.Range('G7').Value}")
            print(f"Highest Risk-Adjusted Opp: {ws_dash.Range('I7').Value}")
            print(f"Recommended Priority: {ws_dash.Range('K7').Value}")
        except Exception as e:
            print(f"Could not read dashboard highlights: {e}")

        # Check Strategic Scorecard results
        try:
            ws_sc = wb.Worksheets("Strategic_Scorecard")
            print("\n--- STRATEGIC SCORECARD RESULTS ---")
            print(f"Winner (Part A): {ws_sc.Range('C30').Value} with lead of {ws_sc.Range('C31').Value} points")
            print(f"Automotive Score: {ws_sc.Range('C26').Value} (Rank {ws_sc.Range('C27').Value})")
            print(f"Industrial Score: {ws_sc.Range('D26').Value} (Rank {ws_sc.Range('D27').Value})")
            print(f"Consumer Score:   {ws_sc.Range('E26').Value} (Rank {ws_sc.Range('E27').Value})")
            print(f"IoT Score:        {ws_sc.Range('F26').Value} (Rank {ws_sc.Range('F27').Value})")
            print(f"Submarket Winner (Part B): {ws_sc.Range('B60').Value} with lead of {ws_sc.Range('B61').Value} points")
        except Exception as e:
            print(f"Could not read scorecard results: {e}")

        # Check Resource Allocation
        try:
            ws_res = wb.Worksheets("Resource_Allocation")
            print("\n--- RESOURCE ALLOCATION ---")
            print(f"Category Total Check: {ws_res.Range('C12').Value}")
            print(f"Application Total Check: {ws_res.Range('D33').Value}")
            print(f"Submarket Total Check: {ws_res.Range('D45').Value}")
            print(f"Robotics Total Resource Allocation: {ws_res.Range('D47').Value}")
        except Exception as e:
            print(f"Could not read resource allocation: {e}")

        # Save to FINAL_ROOT and FINAL_DELIV
        abs_root = os.path.abspath(FINAL_ROOT)
        abs_deliv = os.path.abspath(FINAL_DELIV)
        print(f"\nSaving recalculated workbook to {abs_root}...")
        wb.SaveAs(abs_root, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.Close(SaveChanges=False)
        print(f"Copying to {abs_deliv}...")
        shutil.copyfile(abs_root, abs_deliv)
        print("Workbook saved successfully in both destinations.")

        if errors_found:
            print(f"\n[WARNING] {len(errors_found)} formula errors detected:")
            for cell, code in errors_found[:20]:
                print(f"  {cell}: code {code}")
            return False
        else:
            print("\n[SUCCESS] 0 formula errors found across entire workbook!")
            return True

    except Exception as e:
        print(f"Error during recalculation: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        excel.Quit()

if __name__ == "__main__":
    success = recalc_and_audit()
    if not success:
        sys.exit(1)

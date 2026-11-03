"""Build the Excel workbook with openpyxl (formulas + charts + validation), then recalculate and re-save via Microsoft Excel
(see recalc_excel.py) so that cached values exist for any viewer."""
from openpyxl import Workbook
from common import *
import xl_part1 as p1
import xl_part2 as p2

ORDER = ["Dashboard", "Raw_Data", "Data_Validation", "Revenue_Analysis", "Profitability_Quality", "Growth_Opportunity",
         "Risk_Analysis", "Strategic_Scorecard", "Forecast_Scenarios", "Resource_Allocation", "Sensitivity", "Sources_Assumptions"]
TAB = {"Dashboard": NAVY, "Raw_Data": "7F7F7F", "Data_Validation": "7F7F7F", "Sources_Assumptions": "7F7F7F"}


def build(path=RAW_BUILD):
    wb = Workbook()
    wb.remove(wb.active)
    ws = {n: wb.create_sheet(n) for n in ORDER}
    for n, s in ws.items():
        s.sheet_properties.tabColor = TAB.get(n, BLUE)
    p2.wb_ref.update(ws)
    p1.build_raw(ws["Raw_Data"])
    p1.build_revenue(ws["Revenue_Analysis"])
    p1.build_profit(ws["Profitability_Quality"])
    p1.build_risk(ws["Risk_Analysis"])
    p1.build_forecast(ws["Forecast_Scenarios"])
    p1.build_growth(ws["Growth_Opportunity"])
    p1.fill_revenue_forecast(ws["Revenue_Analysis"])
    p2.build_scorecard(ws["Strategic_Scorecard"])
    p2.build_resources(ws["Resource_Allocation"])
    p2.build_sensitivity(ws["Sensitivity"])
    p2.build_validation(ws["Data_Validation"])
    p2.build_sources(ws["Sources_Assumptions"])
    p2.build_dashboard(ws["Dashboard"])
    for s in ws.values():
        s.sheet_view.zoomScale = 90
        s.page_setup.orientation = "landscape"
        s.page_setup.fitToWidth = 1
        s.page_setup.fitToHeight = 0
        s.sheet_properties.pageSetUpPr.fitToPage = True if s.sheet_properties.pageSetUpPr else None
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    EXPORTS.mkdir(exist_ok=True)
    wb.save(path)
    print("saved", path)


if __name__ == "__main__":
    build()

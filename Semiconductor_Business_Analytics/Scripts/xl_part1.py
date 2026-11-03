"""Workbook sheets part 1: Raw_Data, Revenue_Analysis, Profitability_Quality, Risk_Analysis, Forecast_Scenarios, Growth_Opportunity.
Every calculation is an Excel formula. R is a shared dict of cell addresses used for cross-sheet links."""
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment
from common import *

R = {}  # shared reference registry
SUBMARKETS = ["AI Data Centers", "Factory Automation", "Robotics & Physical AI", "Energy & Smart Grids",
              "Industrial IoT", "Building Automation", "Test & Measurement"]
SUB_CRIT = [  # criterion, report weight, scores for 7 submarkets (research report section 14 scorecard)
    ("Revenue growth", .15, [9, 6, 9, 5, 7, 5, 4]),
    ("Addressable market", .15, [8, 9, 8, 6, 8, 6, 4]),
    ("Competitive advantage", .15, [6, 5, 9, 3, 7, 6, 3]),
    ("Profitability", .15, [7, 8, 8, 5, 6, 5, 9]),
    ("Speed of conversion", .10, [8, 4, 6, 4, 7, 7, 5]),
    ("Design-win visibility", .10, [7, 8, 8, 6, 6, 6, 8]),
    ("Strategic alignment", .10, [7, 8, 10, 4, 8, 6, 4]),
    ("Risk manageability (high = lower risk)", .05, [5, 7, 6, 4, 7, 8, 7]),
    ("Investment feasibility", .05, [8, 6, 8, 3, 8, 8, 4]),
]
SUB_REPORTED_TOTAL = [7.35, 6.60, 8.40, 4.60, 7.15, 5.95, 5.35]


# --------------------------------------------------------------------------------------------- RAW DATA
def build_raw(ws):
    title(ws, "RAW DATA - Research-derived dataset (no separate assignment dataset was supplied)",
          "Source: 'Renesas Industrial Strategy Analysis' research report (research cutoff 2 Nov 2026). Values are as presented in the report; NOT independently re-verified.", 9)
    widths(ws, [8, 54, 14, 14, 14, 16, 18, 30, 62])
    ws.freeze_panes = "C5"
    text_rows(ws, 3, ["IMPORTANT: The original assignment dataset was NOT provided. This sheet is a research-derived dataset. 'Reported*' = figure as cited by the research report from Renesas filings (not re-checked). "
                      "No synthetic observations are stored on this sheet; analyst assumptions live on the analysis sheets (yellow cells)."], 1)
    ws["A3"].font = Font(name=FONT, size=10, bold=True, color=DRED)
    section(ws, 5, "TABLE 1 - Segment financials, JPY billion (report section 7, PDF p.6-7; section 4 table, PDF p.3)", 9)
    header(ws, 6, ["ID", "Metric", "FY2023 (A)", "FY2024 (A)", "FY2025 (A)", "FY2026E (annualised est.)", "Unit", "Data type", "Source"])
    t1 = [("T1.1", "Total consolidated revenue", [1469.4, 1348.4, 1321.2, 1510.4], "JPY bn", "bn"),
          ("T1.2", "Automotive segment revenue", [702.8, 710.0, 639.7, 730.2], "JPY bn", "bn"),
          ("T1.3", "Industrial/Infrastructure/IoT (I&I) segment revenue", [764.7, 636.8, 671.8, 772.0], "JPY bn", "bn"),
          ("T1.4", "Non-GAAP gross margin (consolidated)", [.570, .561, .576, .581], "%", "pct"),
          ("T1.5", "Non-GAAP operating margin (consolidated)", [.341, .295, .293, .327], "%", "pct")]
    for i, (id_, m, vals, u, f) in enumerate(t1):
        r = 7 + i
        W(ws, f"A{r}", id_, "grey"); W(ws, f"B{r}", m, "label")
        for j, v in enumerate(vals):
            W(ws, f"{CL(3 + j)}{r}", v, "src", f)
        W(ws, f"G{r}", u, "grey")
        W(ws, f"H{r}", "Reported* (FY2023-25); Annualised estimate (FY2026E)", "grey", wrap=True)
        W(ws, f"I{r}", "Research report s.7 table; refs [5],[31],[32],[33]", "grey")
        ws.row_dimensions[r].height = 28
    R["raw_rows"] = dict(total=7, auto=8, ii=9, gm=10, om=11, other=16)
    text_rows(ws, 12, ["FY2026E is an ANNUALISED ESTIMATE derived by the report from 1H-2026 momentum and adjusted for the planned Timing divestment - it is not a reported full-year figure."])
    ws["A12"].font = KIND["note"]["font"]
    section(ws, 14, "CALCULATED FIELDS (derived by formula - NOT source data)", 9)
    header(ws, 15, ["ID", "Metric", "FY2023 (A)", "FY2024 (A)", "FY2025 (A)", "FY2026E (est.)", "Unit", "Data type", "Method"])
    W(ws, "A16", "C1.1", "grey"); W(ws, "B16", "Other / unallocated (Total - Automotive - I&I)", "label")
    W(ws, "A17", "C1.2", "grey"); W(ws, "B17", "Other as % of total revenue", "label")
    for j in range(4):
        c = CL(3 + j)
        W(ws, f"{c}16", f"={c}7-{c}8-{c}9", "calc", "bn")
        W(ws, f"{c}17", f"={c}16/{c}7", "calc", "pct")
    for r, u in ((16, "JPY bn"), (17, "%")):
        W(ws, f"G{r}", u, "grey"); W(ws, f"H{r}", "Calculated", "grey")
    W(ws, "I16", "Residual reconciling item; the report does not itemise it", "grey")
    W(ws, "I17", "Other / Total", "grey")

    section(ws, 19, "TABLE 2 - Key values (report sections 4, 6-21). Type: Reported* / Report estimate / Forecast (report) / Third-party market estimate (via report)", 9)
    header(ws, 20, ["ID", "Metric", "Value", "Unit", "Period", "Data type", "Report section / PDF page", "Works-cited ref", "Notes"])
    kv = [  # id, metric, value, unit, period, dtype, section, ref, note, fmt
        ("K01", "I&I non-GAAP gross margin", .606, "%", "FY2025", "Reported*", "s.9, PDF p.8", "[39]", "Segment-level; definitions not verified vs filings", "pct"),
        ("K02", "I&I non-GAAP operating margin", .272, "%", "FY2025", "Reported*", "s.9, PDF p.8", "[39]", "Lower than consolidated 29.3% (see Data_Validation F-04)", "pct"),
        ("K03", "Pure Industrial share of I&I - LOW", .40, "%", "n/a", "Report estimate", "s.4, PDF p.3", "-", "Analyst-style assumption by report; NOT a reported figure", "pct"),
        ("K04", "Pure Industrial share of I&I - HIGH", .45, "%", "n/a", "Report estimate", "s.4, PDF p.3", "-", "", "pct"),
        ("K05", "Pure Industrial share of I&I - POINT estimate", .425, "%", "FY2025", "Report estimate", "s.4 table, PDF p.3", "-", "Used for 285.5; only FY2025 share is given", "pct"),
        ("K06", "Report-stated pure Industrial FY2025 - low end", 282.0, "JPY bn", "FY2025", "Report estimate", "s.4 text, PDF p.3", "-", "Inconsistent: 40% x 671.8 = 268.7 (F-03)", "bn"),
        ("K07", "Report-stated pure Industrial FY2025 - high end", 302.0, "JPY bn", "FY2025", "Report estimate", "s.4 text, PDF p.3", "-", "45% x 671.8 = 302.3 (consistent)", "bn"),
        ("K08", "Report point estimate pure Industrial FY2025", 285.5, "JPY bn", "FY2025", "Report estimate", "s.4 table, PDF p.3", "-", "42.5% x 671.8 = 285.5 (consistent)", "bn"),
        ("K09", "Estimated pure Industrial revenue FY2026", 320.0, "JPY bn", "FY2026E", "Report estimate", "s.17, PDF p.14", "-", "Report baseline for the FY2027 forecast", "bn"),
        ("K10", "FY2027 pure Industrial revenue - Base", 365.0, "JPY bn", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "Research-based scenario, not a guarantee", "bn"),
        ("K11", "FY2027 pure Industrial revenue - Bull", 410.0, "JPY bn", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "bn"),
        ("K12", "FY2027 pure Industrial revenue - Bear", 295.0, "JPY bn", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "bn"),
        ("K13", "Gross margin proxy - Base", .615, "%", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "pct"),
        ("K14", "Gross margin proxy - Bull", .635, "%", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "pct"),
        ("K15", "Gross margin proxy - Bear", .55, "%", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "pct"),
        ("K16", "Scenario probability - Base", .60, "%", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "pct"),
        ("K17", "Scenario probability - Bull", .20, "%", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "pct"),
        ("K18", "Scenario probability - Bear", .20, "%", "FY2027", "Forecast (report)", "s.17, PDF p.14", "-", "", "pct"),
        ("K19", "Corporate gross-margin baseline for KPI", .58, "%", "n/a", "Reported*", "s.21, PDF p.18", "-", "Robotics GM KPI threshold", "pct"),
        ("K20", "Attach-rate KPI target (analog/power ICs per processing node)", 2.5, "ratio", "n/a", "Report target", "s.21, PDF p.18", "-", "Management KPI target, not an observed value", "sc"),
        ("M01", "Factory automation market - 2025", 312.0, "USD bn", "2025", "Third-party market estimate", "s.6, PDF p.4-5", "[23],[24]", "Total market, not semiconductor content", "bn"),
        ("M02", "Factory automation market - 2030", 435.0, "USD bn", "2030", "Third-party market estimate", "s.6, PDF p.4-5", "[23]", "", "bn"),
        ("M03", "AI infrastructure market - 2024", 132.0, "USD bn", "2024", "Third-party market estimate", "s.6, PDF p.5", "[27]", "Total market, not semiconductor content", "bn"),
        ("M04", "AI infrastructure market - 2030", 371.0, "USD bn", "2030", "Third-party market estimate", "s.6, PDF p.5", "[27]", "", "bn"),
        ("M05", "Physical AI / robotics market - 2035 ('over')", 58.0, "USD bn", "2035", "Third-party market estimate", "s.6, PDF p.5", "[26]", "No base-year size given -> no CAGR computable; total market, not Renesas SAM", "bn"),
        ("M06", "Robotics & Physical AI market CAGR (stated, 's.6')", .30, "%", "to 2035", "Third-party market estimate", "s.6, PDF p.5", "[26]", "'exceeding 30%' - conflicts with M07 (F-07)", "pct"),
        ("M07", "Edge-AI / industrial robotics semiconductor CAGR (stated, 's.11')", .25, "%", "next decade", "Third-party market estimate", "s.11, PDF p.9", "[26]", "'exceeding 25%'", "pct"),
        ("M08", "Analog IC market - 2025", 89.5, "USD bn", "2025", "Third-party market estimate", "s.11, PDF p.9", "[30]", "", "bn"),
        ("M09", "Analog IC market CAGR (stated)", .03, "%", "2025-2035", "Third-party market estimate", "s.11, PDF p.9", "[30]", "", "pct"),
        ("M10", "Analog IC market - 2035", 120.0, "USD bn", "2035", "Third-party market estimate", "s.11, PDF p.9", "[30]", "", "bn"),
        ("M11", "Power semiconductor market", 46.1, "USD bn", "n/s", "Third-party market estimate", "s.11, PDF p.9", "[25]", "Year not stated in report", "bn"),
        ("M12", "GaN power semiconductor CAGR (stated)", .11, "%", "n/s", "Third-party market estimate", "s.11, PDF p.9", "[25]", "'over 11%'", "pct"),
        ("C01", "Altium revenue (pre-acquisition)", 263.0, "USD m", "pre-deal", "Reported*", "s.3, PDF p.2", "[1]", "~3% of Renesas revenue: too small to alter revenue quality (F-09)", "bn"),
        ("C02", "Altium recurring revenue share", .77, "%", "pre-deal", "Reported*", "s.3, PDF p.2", "[1]", "", "pct"),
        ("C03", "Altium acquisition price", 5.9, "USD bn", "2024", "Reported*", "s.1, PDF p.1", "[1],[2]", "", "bn"),
        ("C04", "Wolfspeed RSA impairment loss (expected)", 250.0, "JPY bn", "FY2025", "Reported*", "s.7, PDF p.7", "[3],[34]", "", "bn"),
        ("C05", "Timing business sale price (to SiTime)", 3.0, "USD bn", "closing 2026", "Reported*", "s.7, PDF p.7", "[36]", "50% cash / 50% stock -> only ~USD 1.5bn is cash (F-10)", "bn"),
        ("C06", "Timing divestment non-recurring gain", 1.5, "USD bn", "closing 2026", "Reported*", "s.7, PDF p.7", "[35]", "", "bn"),
        ("C07", "Timing divestment non-recurring gain", 234.0, "JPY bn", "closing 2026", "Reported*", "s.7, PDF p.7", "[35]", "", "bn"),
        ("C08", "Maximum single external customer share of revenue (stated cap)", .10, "%", "current", "Reported*", "s.8, PDF p.8", "[37]", "'no single customer above 10%'", "pct"),
        ("C09", "China share of revenue (historical, board question)", .30, "%", "historical", "Report claim", "s.22 Q8 / s.19", "[15]", "Weakly sourced (aggregator) - treat with caution (F-11)", "pct"),
        ("C10", "Industrial design-win qualification - low", 12, "months", "n/a", "Report estimate", "s.3, PDF p.2", "-", "'12 to 24 months'", "int"),
        ("C11", "Industrial design-win qualification - high", 24, "months", "n/a", "Report estimate", "s.3, PDF p.2", "-", "", "int"),
        ("C12", "Factory automation qualification (up to)", 36, "months", "n/a", "Report estimate", "s.6, PDF p.5", "-", "", "int"),
        ("C13", "Robotics design-win timeline - low", 12, "months", "n/a", "Report estimate", "s.6, PDF p.5", "-", "'12 to 18 months'", "int"),
        ("C14", "Robotics design-win timeline - high", 18, "months", "n/a", "Report estimate", "s.6, PDF p.5", "-", "", "int"),
        ("C15", "China analog suppliers' PMIC undercut (low)", .30, "%", "current", "Report claim", "s.18, PDF p.15-16", "[29]", "'30-40%'", "pct"),
        ("C16", "China analog suppliers' PMIC undercut (high)", .40, "%", "current", "Report claim", "s.18, PDF p.15-16", "[29]", "", "pct"),
        ("C17", "Texas Instruments gross margin", .614, "%", "Q2 2026", "Third-party", "s.10, PDF p.9", "[40]", "'routinely exceeding'", "pct"),
        ("C18", "Analog Devices gross margin", .63, "%", "FY2025", "Third-party", "s.10, PDF p.9", "[41]", "'exceeding'", "pct"),
        ("C19", "Test & measurement gross margins (stated)", .65, "%", "n/a", "Third-party", "s.6, PDF p.6", "[29]", "'routinely exceed'", "pct"),
        ("C20", "Industrial product lifecycle - low", 7, "years", "n/a", "Report estimate", "s.3, PDF p.2", "-", "'seven to fifteen years'", "int"),
        ("C21", "Industrial product lifecycle - high", 15, "years", "n/a", "Report estimate", "s.3, PDF p.2", "-", "", "int"),
    ]
    r = 21
    for (id_, m, v, u, per, dt, sec, ref, note, f) in kv:
        W(ws, f"A{r}", id_, "grey"); W(ws, f"B{r}", m, "label", wrap=True)
        W(ws, f"C{r}", v, "src", f)
        W(ws, f"D{r}", u, "grey"); W(ws, f"E{r}", per, "grey"); W(ws, f"F{r}", dt, "grey", wrap=True)
        W(ws, f"G{r}", sec, "grey"); W(ws, f"H{r}", ref, "grey"); W(ws, f"I{r}", note, "grey", wrap=True)
        ws.row_dimensions[r].height = 28 if (len(m) > 50 or len(note) > 60 or len(dt) > 22) else 16
        R[id_] = f"Raw_Data!$C${r}"
        r += 1
    last_kv = r - 1
    R["raw_ids_range"] = f"Raw_Data!$A$21:$A${last_kv}"
    R["raw_vals_range"] = f"Raw_Data!$C$21:$C${last_kv}"
    W(ws, f"A{r}", "C2.1", "grey"); W(ws, f"B{r}", "Implied USD/JPY (= C07 / C06) - calculated, used only for Altium size context", "label", wrap=True)
    W(ws, f"C{r}", f"=C{r-14}/C{r-15}", "calc", "0.0")  # placeholder, fixed below
    ws[f"C{r}"].value = f"={R['C07'].split('!')[1]}/{R['C06'].split('!')[1]}"
    W(ws, f"D{r}", "JPY per USD", "grey"); W(ws, f"E{r}", "2026", "grey"); W(ws, f"F{r}", "Calculated", "grey")
    ws.row_dimensions[r].height = 28
    R["FX"] = f"Raw_Data!$C${r}"
    r += 2
    section(ws, r, "TABLE 3 - Research report weighted commercial scorecard for Industrial submarkets (s.14, PDF p.11-12) - SOURCE INPUTS (analyst-style scores by the report)", 10)
    r += 1
    header(ws, r, ["ID", "Criterion", "Report weight"] + SUBMARKETS)
    ws.row_dimensions[r].height = 42
    r0 = r + 1
    for i, (crit, w, sc) in enumerate(SUB_CRIT):
        rr = r0 + i
        W(ws, f"A{rr}", f"S{i+1}", "grey"); W(ws, f"B{rr}", crit, "label")
        W(ws, f"C{rr}", w, "src", "pct")
        for j, s in enumerate(sc):
            W(ws, f"{CL(4 + j)}{rr}", s, "src", "int")
    rr = r0 + len(SUB_CRIT)
    W(ws, f"A{rr}", "S10", "grey"); W(ws, f"B{rr}", "Weighted total AS REPORTED in the report", "label", bold=True)
    W(ws, f"C{rr}", f"=SUM(C{r0}:C{rr-1})", "calc", "pct")
    for j, t in enumerate(SUB_REPORTED_TOTAL):
        W(ws, f"{CL(4 + j)}{rr}", t, "src", "0.00")
    R["sub_first_row"], R["sub_last_row"], R["sub_total_row"] = r0, rr - 1, rr
    text_rows(ws, rr + 1, ["Note: the report-stated totals do not all reconcile to its own scores and weights (see Strategic_Scorecard Part B and Data_Validation F-05)."])
    ws[f"A{rr+1}"].font = KIND["note"]["font"]
    ws.conditional_formatting.add(f"C{r0}:C{rr}", CellIsRule(operator="notEqual", formula=["1"], font=Font(color=DRED)) if False else CellIsRule(operator="lessThan", formula=["0"], font=Font(color=DRED)))


def rawy(key, yi):
    row = R["raw_rows"][key]
    return f"Raw_Data!${CL(3 + yi)}${row}"


# --------------------------------------------------------------------------------------------- REVENUE
def build_revenue(ws):
    title(ws, "REVENUE ANALYSIS - Automotive | Industrial | Consumer | IoT (JPY billion)",
          "Automotive & I&I are REPORTED segments. Industrial = ESTIMATE (42.5% of I&I). Consumer/IoT = ILLUSTRATIVE split of the non-industrial remainder (no source). FY2026E = annualised estimate.", 9)
    widths(ws, [50, 14, 14, 14, 14, 16, 18, 16, 62])
    ws.freeze_panes = "B4"
    section(ws, 4, "1. Revenue by application (JPY bn)", 9)
    header(ws, 5, ["Application"] + YEARS + ["CAGR FY23-25 (actuals only)", "CAGR FY23-26E (incl. estimate)", "Change FY25 -> FY26E", "Data type / basis"], h=42)
    labels = [("Automotive", "Reported segment (Raw_Data T1.2)"),
              ("Industrial (ESTIMATE)", "ESTIMATE: I&I x pure-Industrial share (FY23-25 share held at FY2025 point estimate - assumption); FY2026E = report estimate K09"),
              ("Consumer (ILLUSTRATIVE)", "ILLUSTRATIVE: (I&I - Industrial) x Consumer share assumption B18 - not a source figure"),
              ("IoT & Infrastructure (ILLUSTRATIVE)", "ILLUSTRATIVE: remainder of non-industrial I&I"),
              ("Other / unallocated", "Calculated residual (Raw_Data C1.1)")]
    for i, (lab, basis) in enumerate(labels):
        r = 6 + i
        W(ws, f"A{r}", lab, "label", bold=True)
        W(ws, f"I{r}", basis, "grey", wrap=True)
        ws.row_dimensions[r].height = 28
        for yi in range(4):
            c = CL(2 + yi)
            if i == 0:
                f = f"={rawy('auto', yi)}"
            elif i == 1:
                f = f"={rawy('ii', yi)}*$B$19" if yi < 3 else f"={R['K09']}"
            elif i == 2:
                f = f"=({c}14-{c}7)*$B$18"
            elif i == 3:
                f = f"=({c}14-{c}7)*(1-$B$18)"
            else:
                f = f"={rawy('other', yi)}"
            W(ws, f"{c}{r}", f, "calc", "bn")
    W(ws, "A11", "Total (sum of components)", "label", bold=True)
    W(ws, "A12", "Reported total revenue (Raw_Data)", "label")
    W(ws, "A13", "Reconciliation difference (must be 0.0)", "label")
    W(ws, "A14", "I&I reported segment (Raw_Data)", "label")
    W(ws, "A15", "Industrial as % of I&I (implied)", "label")
    for yi in range(4):
        c = CL(2 + yi)
        W(ws, f"{c}11", f"=SUM({c}6:{c}10)", "out", "bn")
        W(ws, f"{c}12", f"={rawy('total', yi)}", "calc", "bn")
        W(ws, f"{c}13", f"=ROUND({c}11-{c}12,6)", "calc", "bn")
        W(ws, f"{c}14", f"={rawy('ii', yi)}", "calc", "bn")
        W(ws, f"{c}15", f"={c}7/{c}14", "calc", "pct")
    for r in (6, 7, 8, 9, 11, 14):
        W(ws, f"F{r}", f"=(D{r}/B{r})^(1/2)-1", "calc", "pct")
        W(ws, f"G{r}", f"=(E{r}/B{r})^(1/3)-1", "calc", "pct")
        W(ws, f"H{r}", f"=E{r}-D{r}", "calc", "bn")
    for c in "FGH":
        W(ws, f"{c}10", "n/m", "grey", h="center")
    W(ws, "I11", "Components always reconcile to reported total (check row 13)", "grey")
    W(ws, "I13", "Avoids double counting: Industrial + Consumer + IoT = I&I exactly", "grey")
    W(ws, "I15", "FY26E share is DERIVED from the report's 320 and annualised I&I 772.0 (41.5%), inside the 40-45% range", "grey", wrap=True)
    ws.row_dimensions[15].height = 28
    section(ws, 17, "Analyst assumptions used on this sheet", 9)
    W(ws, "A18", "Consumer share of non-industrial I&I remainder (ILLUSTRATIVE)", "label")
    W(ws, "B18", 0.40, "assump", "pct")
    W(ws, "C18", "No source splits Consumer vs IoT/Infrastructure. Report only says the remainder mixes consumer electronics, smart home and hyperscale data-centre infrastructure. Change to test.", "note", border=False)
    W(ws, "A19", "Pure Industrial share of I&I applied to FY2023-FY2025 (point estimate)", "label")
    W(ws, "B19", f"={R['K05']}", "calc", "pct")
    W(ws, "C19", "Report supplies the share for FY2025 only; holding it constant makes Industrial/Consumer/IoT growth identical to I&I growth FY23-25 BY CONSTRUCTION.", "note", border=False)
    R["rev_cons_share"] = "Revenue_Analysis!$B$18"
    R["rev_rows"] = dict(Automotive=6, Industrial=7, Consumer=8, IoT=9, Other=10, Total=11, II=14)

    section(ws, 21, "2. Revenue contribution (% of total)", 9)
    header(ws, 22, ["Application"] + YEARS)
    for i in range(5):
        r = 23 + i
        W(ws, f"A{r}", f"=A{6+i}", "label")
        for yi in range(4):
            c = CL(2 + yi)
            W(ws, f"{c}{r}", f"={c}{6+i}/{c}$11", "calc", "pct")
    W(ws, "A28", "Total", "label", bold=True)
    for yi in range(4):
        c = CL(2 + yi)
        W(ws, f"{c}28", f"=SUM({c}23:{c}27)", "out", "pct")
    section(ws, 30, "3. Year-over-year growth = Current / Previous - 1  (FY2026E column is vs an ESTIMATE - label accordingly)", 9)
    header(ws, 31, ["Application"] + YEARS)
    yoy_labels = ["Automotive", "Industrial (ESTIMATE)", "Consumer (ILLUSTRATIVE)", "IoT & Infrastructure (ILLUSTRATIVE)", "Total revenue", "I&I reported segment"]
    yoy_src = [6, 7, 8, 9, 11, 14]
    for i, (lab, sr) in enumerate(zip(yoy_labels, yoy_src)):
        r = 32 + i
        W(ws, f"A{r}", lab, "label")
        W(ws, f"B{r}", "n/a", "grey", h="center")
        for yi in range(1, 4):
            c, p = CL(2 + yi), CL(1 + yi)
            W(ws, f"{c}{r}", f"={c}{sr}/{p}{sr}-1", "calc", "pct")
    ws.conditional_formatting.add("C32:E37", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", bgColor=RED), font=Font(color=DRED)))
    ws.conditional_formatting.add("C32:E37", CellIsRule(operator="greaterThan", formula=["0.05"], fill=PatternFill("solid", bgColor=GREEN)))
    section(ws, 39, "4. Absolute revenue change (JPY bn)", 9)
    header(ws, 40, ["Application"] + YEARS)
    for i in range(5):
        r = 41 + i
        W(ws, f"A{r}", f"=A{6+i}", "label")
        W(ws, f"B{r}", "n/a", "grey", h="center")
        for yi in range(1, 4):
            c, p = CL(2 + yi), CL(1 + yi)
            W(ws, f"{c}{r}", f"={c}{6+i}-{p}{6+i}", "calc", "bn")
    W(ws, "A46", "Total revenue", "label", bold=True)
    W(ws, "B46", "n/a", "grey", h="center")
    for yi in range(1, 4):
        c, p = CL(2 + yi), CL(1 + yi)
        W(ws, f"{c}46", f"={c}11-{p}11", "out", "bn")
    section(ws, 48, "5. Revenue concentration", 9)
    header(ws, 49, ["Measure"] + YEARS)
    W(ws, "A50", "HHI across the five revenue buckets (0-10,000; >2,500 = highly concentrated)", "label")
    W(ws, "A51", "Largest bucket share", "label"); W(ws, "A52", "Largest bucket", "label")
    W(ws, "A53", "Automotive + Industrial share of total", "label")
    for yi in range(4):
        c = CL(2 + yi)
        W(ws, f"{c}50", f"=SUMPRODUCT({c}23:{c}27,{c}23:{c}27)*10000", "calc", "#,##0")
        W(ws, f"{c}51", f"=MAX({c}23:{c}27)", "calc", "pct")
        W(ws, f"{c}52", f"=INDEX($A$6:$A$10,MATCH({c}51,{c}23:{c}27,0))", "calc", h="center")
        W(ws, f"{c}53", f"={c}23+{c}24", "calc", "pct")
    ws.row_dimensions[52].height = 30
    for c in "BCDE":
        ws[f"{c}52"].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    section(ws, 55, "6. Industrial disaggregation methodology (FY2025) - Industrial is an ESTIMATE, never the whole I&I segment", 9)
    header(ws, 56, ["Case", "Share of I&I", "I&I FY2025 (JPY bn)", "Pure Industrial (calc)", "Report-stated", "Calc - report", "Check", "", "Comment"], h=32)
    cases = [("Low (40%)", "K03", "K06", "Report states 282 as the low end, but 40% x 671.8 = 268.7 -> the report's low end looks overstated by ~13 bn"),
             ("Point (42.5%)", "K05", "K08", "Matches the report table (285.5)"),
             ("High (45%)", "K04", "K07", "Matches the report text (302)")]
    for i, (lab, share, rep, cm) in enumerate(cases):
        r = 57 + i
        W(ws, f"A{r}", lab, "label")
        W(ws, f"B{r}", f"={R[share]}", "calc", "pct")
        W(ws, f"C{r}", "=$D$14", "calc", "bn")
        W(ws, f"D{r}", f"=B{r}*C{r}", "out", "bn")
        W(ws, f"E{r}", f"={R[rep]}", "calc", "bn")
        W(ws, f"F{r}", f"=D{r}-E{r}", "calc", "bn")
        W(ws, f"G{r}", f'=IF(ABS(F{r})<=1,"OK","FLAG")', "calc", h="center")
        W(ws, f"I{r}", cm, "grey", wrap=True)
        ws.row_dimensions[r].height = 28
    ws.conditional_formatting.add("G57:G59", CellIsRule(operator="equal", formula=['"FLAG"'], fill=PatternFill("solid", bgColor=RED), font=Font(color=DRED, bold=True)))
    ws.conditional_formatting.add("G57:G59", CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", bgColor=GREEN)))
    W(ws, "A60", "Uncertainty band on FY2025 Industrial revenue (High - Low, JPY bn)", "label")
    W(ws, "D60", "=D59-D57", "calc", "bn")
    W(ws, "A61", "Band as % of point estimate", "label")
    W(ws, "D61", "=D60/D58", "calc", "pct")
    R["rev_band_pct"] = "Revenue_Analysis!$D$61"
    # section 7 forecast link rows 63+ (filled later)
    section(ws, 63, "7. FY2027 forecast revenue by application (probability-weighted; links to Forecast_Scenarios)", 9)
    header(ws, 64, ["Application", "FY2026E (est.)", "FY2027E expected", "Expected growth", "Change (JPY bn)"])
    section(ws, 72, "8. Limitations - read before using growth rates", 9)
    notes = ["- Only two segments are reported (Automotive; I&I). Industrial, Consumer and IoT revenues are NOT reported; Industrial is a report estimate and Consumer/IoT split is illustrative.",
             "- Because the Industrial share is held constant for FY2023-25, Industrial/Consumer/IoT historical growth equals I&I growth by construction: it cannot discriminate between those three applications.",
             "- FY2026E is an annualised estimate; growth into FY2026E (+14%) mixes a rebound with the report's own estimation method. CAGR FY23-25 uses actuals only (2 periods).",
             "- Total-vs-segment reconciliation leaves a small 'Other' residual (0.5%-0.7% of revenue); it is shown explicitly rather than silently allocated.",
             "- Revenue contribution and growth are NOT sufficient to rank applications: see Profitability_Quality, Risk_Analysis and Strategic_Scorecard."]
    text_rows(ws, 73, notes)
    # charts
    ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100
    ch.title = "Revenue by application (JPY bn; FY26E = estimate)"
    ch.y_axis.title = "JPY bn"
    data = Reference(ws, min_col=1, max_col=5, min_row=6, max_row=10)
    ch.add_data(data, from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=5, min_row=5))
    ch.height, ch.width = 9, 17
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ws.add_chart(ch, "K4")
    lc = LineChart(); lc.title = "YoY growth by application (FY26E vs estimate)"
    lc.add_data(Reference(ws, min_col=1, max_col=5, min_row=32, max_row=36), from_rows=True, titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=2, max_col=5, min_row=31))
    lc.y_axis.number_format = "0%"; lc.height, lc.width = 9, 17
    lc.x_axis.delete = False; lc.y_axis.delete = False
    ws.add_chart(lc, "K24")


def fill_revenue_forecast(ws):
    for i, app in enumerate(APPS):
        r = 65 + i
        W(ws, f"A{r}", f"=A{6+i}", "label")
        W(ws, f"B{r}", f"=E{6+i}", "calc", "bn")
        W(ws, f"C{r}", f"={R['fc_sum_rev'][app]}", "calc", "bn")
        W(ws, f"D{r}", f"=C{r}/B{r}-1", "calc", "pct")
        W(ws, f"E{r}", f"=C{r}-B{r}", "calc", "bn")
    W(ws, "A69", "Other / unallocated (held flat)", "label")
    W(ws, "B69", "=E10", "calc", "bn"); W(ws, "C69", f"={R['fc_other']}", "calc", "bn")
    W(ws, "D69", "=C69/B69-1", "calc", "pct"); W(ws, "E69", "=C69-B69", "calc", "bn")
    W(ws, "A70", "Total company (expected)", "label", bold=True)
    W(ws, "B70", "=SUM(B65:B69)", "out", "bn"); W(ws, "C70", "=SUM(C65:C69)", "out", "bn")
    W(ws, "D70", "=C70/B70-1", "out", "pct"); W(ws, "E70", "=C70-B70", "out", "bn")
    R["rev_total_fc"] = "Revenue_Analysis!$C$70"
    R["rev_total_fc_g"] = "Revenue_Analysis!$D$70"


# --------------------------------------------------------------------------------------------- PROFITABILITY
QUALITY = [  # factor, weight, [auto, ind, cons, iot], evidence
    ("Revenue stability (volatility of YoY growth)", .15, [6, 5, 3, 4], "Measured proxy below: I&I YoY volatility is HIGHER than Automotive over FY24-26E, contradicting the report's 'resilience' claim; Consumer most cyclical (analyst)."),
    ("Customer concentration (high = more diversified)", .15, [4, 8, 4, 7], "Report s.8: tier-1 auto customers wield monopsony power; Industrial base fragmented; Consumer relies on few OEMs/platforms (analyst)."),
    ("Recurring vs transactional revenue", .10, [5, 6, 2, 4], "Only Altium (USD 263m, 77% recurring) is recurring - ~3% of group revenue; silicon is transactional but multi-year platform."),
    ("Customer retention / stickiness", .10, [7, 7, 3, 5], "Qualified designs persist for the product life; consumer designs re-bid each cycle (analyst)."),
    ("Product lifecycle length", .10, [8, 9, 2, 5], "Report s.3: industrial lifecycles 7-15 years; automotive platforms long; consumer short (analyst)."),
    ("Pricing power", .15, [4, 6, 2, 5], "Auto: continuous price-downs + Chinese entrants (report s.8, s.18); Industrial: Winning Combinations raise switching costs but TI/Infineon/ADI compete."),
    ("Switching costs", .15, [7, 8, 2, 4], "Re-qualification effort; PMIC+MCU bundles raise cost of modular substitution (report s.8). Auto = AEC/ISO re-qualification."),
    ("Capital intensity (high = lighter)", .10, [4, 7, 7, 7], "Report s.9: fab-lite model; mature-node auto fabs in Japan; SiC retreat caused 250bn JPY impairment (auto power)."),
]


def build_profit(ws):
    title(ws, "PROFITABILITY & REVENUE QUALITY",
          "Application-level profit data is NOT disclosed. Automotive margins are DERIVED; Industrial/Consumer/IoT margins are PROXIES that reconcile exactly to the reported I&I segment margin.", 8)
    widths(ws, [52, 14, 14, 14, 14, 14, 70, 4])
    ws.freeze_panes = "B4"
    section(ws, 4, "1. Reported margins (non-GAAP) - consolidated and I&I", 8)
    header(ws, 5, ["Metric"] + YEARS + ["", "Data type"])
    W(ws, "A6", "Consolidated gross margin", "label"); W(ws, "A7", "Consolidated operating margin", "label")
    W(ws, "A8", "I&I segment gross margin", "label"); W(ws, "A9", "I&I segment operating margin", "label")
    for yi in range(4):
        c = CL(2 + yi)
        W(ws, f"{c}6", f"={rawy('gm', yi)}", "calc", "pct"); W(ws, f"{c}7", f"={rawy('om', yi)}", "calc", "pct")
        if yi == 2:
            W(ws, f"{c}8", f"={R['K01']}", "calc", "pct"); W(ws, f"{c}9", f"={R['K02']}", "calc", "pct")
        else:
            W(ws, f"{c}8", "n/d", "grey", h="center"); W(ws, f"{c}9", "n/d", "grey", h="center")
    W(ws, "G6", "Reported* (FY26E annualised est.)", "grey"); W(ws, "G7", "Reported* (FY26E annualised est.)", "grey")
    W(ws, "G8", "Reported* FY2025 only; other years not disclosed in report (n/d)", "grey"); W(ws, "G9", "Reported* FY2025 only", "grey")
    section(ws, 11, "2. Implied Automotive margin FY2025 (DERIVED = consolidated less I&I; 'Other' assumed at consolidated margin)", 8)
    items = [("Consolidated revenue FY2025", f"={rawy('total', 2)}", "bn"), ("Automotive revenue FY2025", f"={rawy('auto', 2)}", "bn"),
             ("I&I revenue FY2025", f"={rawy('ii', 2)}", "bn"), ("Other / unallocated FY2025", f"={rawy('other', 2)}", "bn")]
    for i, (l, f, fm) in enumerate(items):
        W(ws, f"A{12+i}", l, "label"); W(ws, f"B{12+i}", f, "calc", fm)
    W(ws, "A16", "Implied Automotive gross margin (derived)", "label")
    W(ws, "B16", "=(D6*B12-D8*B14-D6*B15)/B13", "out", "pct")
    W(ws, "A17", "Implied Automotive operating margin (derived)", "label")
    W(ws, "B17", "=(D7*B12-D9*B14-D7*B15)/B13", "out", "pct")
    W(ws, "A18", "Gross-margin gap: I&I less Automotive (pp)", "label"); W(ws, "B18", "=D8-B16", "calc", "pp")
    W(ws, "A19", "Operating-margin gap: I&I less Automotive (pp)", "label"); W(ws, "B19", "=D9-B17", "calc", "pp")
    W(ws, "G16", "DERIVED - not reported. Segment-margin definitions (e.g. unallocated corporate costs) are unverified, so treat as indicative.", "note", border=False)
    W(ws, "G18", "I&I has higher GROSS margin ...", "note", border=False)
    W(ws, "G19", "... but LOWER operating margin than Automotive (implied): 'superior structural profitability' is only true at gross-margin level (F-04).", "note", border=False)
    ws.conditional_formatting.add("B18:B19", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=DRED, bold=True)))
    R["auto_gm"], R["auto_om"] = "Profitability_Quality!$B$16", "Profitability_Quality!$B$17"

    section(ws, 21, "3. Application profitability proxies FY2025 (reconcile exactly to I&I segment)", 8)
    header(ws, 22, ["Metric"] + APPS + ["Check / total", "Basis"])
    W(ws, "A23", "Revenue FY2025 (JPY bn)", "label")
    for i, app in enumerate(APPS):
        W(ws, f"{CL(2+i)}23", f"=Revenue_Analysis!$D${R['rev_rows'][app]}", "calc", "bn")
    W(ws, "F23", "=SUM(B23:E23)", "calc", "bn")
    W(ws, "A24", "GM premium of Industrial over I&I average (pp) - ANALYST ASSUMPTION", "label"); W(ws, "C24", 0.02, "assump", "pp")
    W(ws, "A25", "GM discount of Consumer vs non-industrial remainder (pp) - ANALYST ASSUMPTION", "label"); W(ws, "D25", -0.04, "assump", "pp")
    W(ws, "A26", "OM premium of Industrial over I&I average (pp) - ANALYST ASSUMPTION", "label"); W(ws, "C26", 0.0, "assump", "pp")
    W(ws, "A27", "OM discount of Consumer vs non-industrial remainder (pp) - ANALYST ASSUMPTION", "label"); W(ws, "D27", -0.02, "assump", "pp")
    W(ws, "A28", "Non-industrial remainder gross margin (derived so I&I reconciles)", "label")
    W(ws, "B28", "=(D8*(C23+D23+E23)-(D8+C24)*C23)/(D23+E23)", "calc", "pct")
    W(ws, "A29", "Non-industrial remainder operating margin (derived)", "label")
    W(ws, "B29", "=(D9*(C23+D23+E23)-(D9+C26)*C23)/(D23+E23)", "calc", "pct")
    W(ws, "A30", "IoT adjustment vs remainder (derived so weights net to zero)", "label")
    W(ws, "B30", "=-D25*D23/E23", "calc", "pp")
    W(ws, "C30", "=-D27*D23/E23", "calc", "pp")
    W(ws, "A31", "GROSS MARGIN - proxy FY2025", "label", bold=True)
    W(ws, "B31", "=B16", "out", "pct"); W(ws, "C31", "=D8+C24", "out", "pct")
    W(ws, "D31", "=B28+D25", "out", "pct"); W(ws, "E31", "=B28+B30", "out", "pct")
    W(ws, "A32", "OPERATING MARGIN - proxy FY2025", "label", bold=True)
    W(ws, "B32", "=B17", "out", "pct"); W(ws, "C32", "=D9+C26", "out", "pct")
    W(ws, "D32", "=B29+D27", "out", "pct"); W(ws, "E32", "=B29+C30", "out", "pct")
    W(ws, "A33", "Gross profit FY2025 (JPY bn)", "label")
    W(ws, "A34", "Operating profit FY2025 (JPY bn)", "label")
    W(ws, "A35", "Share of the four applications' gross profit", "label")
    for i in range(4):
        c = CL(2 + i)
        W(ws, f"{c}33", f"={c}23*{c}31", "calc", "bn"); W(ws, f"{c}34", f"={c}23*{c}32", "calc", "bn")
        W(ws, f"{c}35", f"={c}33/$F$33", "calc", "pct")
    W(ws, "F33", "=SUM(B33:E33)", "calc", "bn"); W(ws, "F34", "=SUM(B34:E34)", "calc", "bn"); W(ws, "F35", "=SUM(B35:E35)", "calc", "pct")
    W(ws, "A36", "Reconciliation: Industrial+Consumer+IoT GP less I&I GP (must be 0.0)", "label")
    W(ws, "F36", "=ROUND(SUM(C33:E33)-D8*D14*0-D8*B14,6)", "calc", "bn")
    ws["F36"].value = "=ROUND(SUM(C33:E33)-D8*B14,6)"
    W(ws, "A37", "Reconciliation: Industrial+Consumer+IoT OP less I&I OP (must be 0.0)", "label")
    W(ws, "F37", "=ROUND(SUM(C34:E34)-D9*B14,6)", "calc", "bn")
    W(ws, "A38", "Data type", "label")
    for i, t in enumerate(["DERIVED", "PROXY", "PROXY", "PROXY"]):
        W(ws, f"{CL(2+i)}38", t, "flag" if t == "PROXY" else "grey", h="center")
    W(ws, "G31", "Automotive derived (section 2). Others = I&I average +/- analyst adjustments that net to zero across the I&I segment. NOT disclosed figures.", "note", border=False)
    W(ws, "G32", "Operating margins are proxies for planning only.", "note", border=False)
    R["gm_proxy"] = {a: f"Profitability_Quality!${CL(2+i)}$31" for i, a in enumerate(APPS)}
    R["om_proxy"] = {a: f"Profitability_Quality!${CL(2+i)}$32" for i, a in enumerate(APPS)}

    section(ws, 40, "4. Unit economics and other data availability", 8)
    header(ws, 41, ["Metric", "Available?", "", "", "", "", "Treatment in this workbook"])
    avail = [("Profit per unit / ASP by application", "No", "Not available - not invented. Attach-rate effect is illustrated with an INDEX model (Growth_Opportunity s.4)."),
             ("Contribution margin by application", "No", "Gross-margin proxy used; flagged."),
             ("Customer concentration", "Partial", "Report: no single customer >10% of group revenue (C08); application-level concentration is qualitative (section 5 scores)."),
             ("Recurring vs transactional", "Partial", "Only Altium (USD 263m, 77% recurring) - about 3% of revenue; silicon is transactional."),
             ("Customer retention", "No", "Qualitative score only (analyst)."),
             ("Capital intensity", "Qualitative", "Fab-lite model per report s.9; SiC retreat cost 250bn JPY (C04).")]
    for i, (m, a, t) in enumerate(avail):
        r = 42 + i
        W(ws, f"A{r}", m, "label"); W(ws, f"B{r}", a, "grey", h="center"); W(ws, f"G{r}", t, "grey", wrap=True)
        ws.row_dimensions[r].height = 28
    section(ws, 49, "5. Revenue-quality scorecard (1-10, higher = BETTER quality) - ANALYST ASSESSMENT with report evidence", 8)
    header(ws, 50, ["Quality factor", "Weight"] + APPS + ["Evidence / basis"], h=30)
    ws.column_dimensions["B"].width = 14
    for i, (f, w, sc, ev) in enumerate(QUALITY):
        r = 51 + i
        W(ws, f"A{r}", f, "label"); W(ws, f"B{r}", w, "assump", "pct")
        for j, s in enumerate(sc):
            W(ws, f"{CL(3+j)}{r}", s, "assump", "int")
        W(ws, f"G{r}", ev, "grey", wrap=True)
        ws.row_dimensions[r].height = 40
    W(ws, "A59", "Weights check (must be 100%)", "label"); W(ws, "B59", "=SUM(B51:B58)", "calc", "pct")
    W(ws, "A60", "WEIGHTED REVENUE-QUALITY SCORE", "label", bold=True)
    for j in range(4):
        c = CL(3 + j)
        W(ws, f"{c}60", f"=SUMPRODUCT($B$51:$B$58,{c}51:{c}58)", "out", "sc")
    W(ws, "A61", "Rank (1 = best quality)", "label")
    for j in range(4):
        c = CL(3 + j)
        W(ws, f"{c}61", f"=RANK({c}60,$C$60:$F$60)", "calc", "int")
    ws.conditional_formatting.add("C51:F58", ColorScaleRule(start_type="num", start_value=1, start_color="F8696B", mid_type="num", mid_value=5.5, mid_color="FFEB84", end_type="num", end_value=10, end_color="63BE7B"))
    R["quality"] = {a: f"Profitability_Quality!${CL(3+i)}$60" for i, a in enumerate(APPS)}
    R["q_wsum"] = "Profitability_Quality!$B$59"
    section(ws, 63, "Measured proxy: revenue volatility = standard deviation of YoY growth FY24-FY26E (3 observations; indicative only)", 8)
    W(ws, "A64", "Automotive (reported)", "label"); W(ws, "B64", "=STDEV(Revenue_Analysis!C32:E32)", "calc", "pct")
    W(ws, "A65", "I&I segment (reported)", "label"); W(ws, "B65", "=STDEV(Revenue_Analysis!C37:E37)", "calc", "pct")
    W(ws, "G64", "I&I fell 16.7% in FY2024 vs Automotive +1.0%: I&I is NOT more stable than Automotive on this evidence.", "note", border=False)
    R["vol_auto"], R["vol_ii"] = "Profitability_Quality!$B$64", "Profitability_Quality!$B$65"
    section(ws, 67, "Why revenue growth alone is insufficient to choose where to invest", 8)
    text_rows(ws, 68, [
        "1. Growth without margin: +1 yen of low-margin consumer revenue adds far less gross profit than +1 yen of Industrial revenue (see gross-profit rows above).",
        "2. Growth without quality: revenue that depends on a few customers or re-bids every cycle is less predictable and deserves a lower multiple.",
        "3. Growth without cost-to-win: Industrial design-wins need 12-36 months and heavy FAE support, so the YEAR-1 return on incremental resources is lower than headline growth suggests.",
        "4. Growth without risk: the highest growth submarkets (AI data centre, robotics) also carry concentration, geopolitical and substitution risk - hence the weighted scorecard."])


# --------------------------------------------------------------------------------------------- RISK
RISK_FACTORS = [  # factor, weight, [auto, ind, cons, iot], rationale
    ("Market cyclicality", .12, [7, 5, 8, 6], "Auto saw -9.9% FY2025; consumer most cyclical; industrial recovers slower but has long lifecycles. (Report s.4, s.7)"),
    ("Customer concentration", .10, [8, 4, 7, 4], "Auto: tier-1 monopsony power (s.8); Consumer: few platform OEMs; Industrial/IoT fragmented. AI data-centre sub-market is concentrated (s.6)."),
    ("Competitive pressure", .10, [7, 6, 8, 7], "Auto: NXP/Infineon/STM + Chinese entrants; Industrial: TI, ADI, Infineon, STM (s.10)."),
    ("Pricing pressure", .10, [8, 5, 8, 7], "Chinese PMIC undercut 30-40% (s.18); auto price-downs; consumer commoditisation."),
    ("Technology substitution", .08, [6, 5, 6, 6], "RISC-V vs Arm Cortex MCUs in edge AI (s.18); SDV central compute shifts in auto (analyst)."),
    ("Geopolitical exposure", .10, [7, 6, 7, 6], "Export controls, China ~30% of revenue claim (s.19, s.22); heavy-industry robotics China exposure."),
    ("Supply-chain dependence", .08, [6, 5, 5, 5], "TSMC for advanced logic; Japanese mature-node fabs for auto (s.9)."),
    ("Capital requirements", .08, [7, 4, 3, 4], "Auto power (SiC) retreat cost 250bn JPY; fab-lite industrial/IoT (s.7, s.9)."),
    ("Inventory-correction exposure", .08, [7, 5, 8, 6], "FY2024 correction hit I&I -16.7%; channel DOI KPI (s.18, s.21)."),
    ("Design-win conversion period", .08, [7, 7, 3, 4], "Industrial 12-24m, factory automation up to 36m, robotics 12-18m (s.3, s.6). Auto multi-year (analyst)."),
    ("Customer qualification barriers", .08, [8, 6, 2, 3], "Functional-safety (ISO 26262 / IEC 61508) and AEC-Q qualification (s.19, analyst)."),
]
RISK_REG = [  # risk, prob, impact, warning, mitigation, effectiveness, source
    ("Geopolitical & export-control exposure", "High", "Severe", "Expanded US/Japan restrictions on shipments to Chinese AI & robotics firms",
     "Weight resources to N. America/Europe/India; in China sell only differentiated products; localise supply where feasible", .30, "Report s.18"),
    ("RISC-V substitution of Arm Cortex MCUs", "Medium", "High", "RISC-V design-win losses in edge-AI MCUs; customer RFQs requesting open ISA",
     "Offer RISC-V parts; keep Altium/Reality AI architecture-agnostic so software revenue is retained", .40, "Report s.18"),
    ("Altium integration failure", "Low", "Severe", "Talent attrition; Renesas libraries not embedded in cloud platform",
     "Multi-year retention packages; integration management office tracking engagement weekly", .50, "Report s.18"),
    ("Chinese pricing pressure (PMIC undercut 30-40%)", "High", "Medium", "Win/loss notes citing price; ASP erosion on discrete PMICs",
     "Bundle PMIC + MCU/MPU in reference designs to obscure unit price and block discrete substitution", .35, "Report s.18"),
    ("Customer inventory corrections", "Medium", "Medium", "Rising Days of Inventory at distributors; stalled sell-through",
     "Move from booking-price to net-selling-price basis for channel visibility; track DOI by product line", .40, "Report s.18"),
    ("Customer concentration", "Low", "Medium", "Any robotics/AI account >10% of Industrial revenue; reliance on a few hyperscaler-linked accounts",
     "Keep robotics funnel broad (500 target accounts); cap AI data-centre dependence; monitor top-10 share", .50, "Analyst (report s.8: no customer >10%)"),
    ("Long design-win conversion cycles", "High", "High", "Evaluation-to-production days rising; pilots not converting",
     "Pre-validated reference designs, Altium-sourced pipeline, subsidised kits; set FY2027 revenue expectation on EXISTING pipeline", .30, "Analyst (report s.3, s.6)"),
]


def build_risk(ws):
    title(ws, "RISK ANALYSIS", "Scoring scale 1-10: HIGHER score = HIGHER risk (this sheet only). The strategic scorecard converts it to 'risk manageability' where higher = better.", 11)
    widths(ws, [44, 12, 12, 12, 12, 12, 56, 62, 14, 14, 26])
    ws.freeze_panes = "B4"
    section(ws, 4, "Methodology", 11)
    text_rows(ws, 5, ["Each application is scored 1 (negligible) to 10 (severe) on 11 risk factors; factor weights are analyst assumptions that sum to 100%. Weighted risk = SUMPRODUCT(weights, scores).",
                      "Scores are ANALYST ASSESSMENTS anchored in the report's qualitative evidence (cited in the 'Basis' column) - they are not measured financial metrics.",
                      "Risk manageability (used in Strategic_Scorecard) = 11 - weighted risk, so that 10 = very manageable and 1 = unmanageable."])
    section(ws, 9, "1. Risk-factor scorecard by application", 11)
    header(ws, 10, ["Risk factor", "Weight"] + APPS + ["Basis for scores (report section)"])
    for i, (f, w, sc, b) in enumerate(RISK_FACTORS):
        r = 11 + i
        W(ws, f"A{r}", f, "label"); W(ws, f"B{r}", w, "assump", "pct")
        for j, s in enumerate(sc):
            W(ws, f"{CL(3+j)}{r}", s, "assump", "int")
        W(ws, f"G{r}", b, "grey", wrap=True)
        ws.row_dimensions[r].height = 30
    W(ws, "A22", "Weights check (must be 100%)", "label"); W(ws, "B22", "=SUM(B11:B21)", "calc", "pct")
    W(ws, "A23", "WEIGHTED RISK SCORE (higher = riskier)", "label", bold=True)
    W(ws, "A24", "Risk band", "label"); W(ws, "A25", "Risk rank (1 = lowest risk)", "label")
    W(ws, "A26", "RISK MANAGEABILITY score (11 - risk; higher = better)", "label", bold=True)
    for j in range(4):
        c = CL(3 + j)
        W(ws, f"{c}23", f"=SUMPRODUCT($B$11:$B$21,{c}11:{c}21)", "out", "sc")
        W(ws, f"{c}24", f'=IF({c}23>=6.5,"High",IF({c}23>=5,"Medium","Low"))', "calc", h="center")
        W(ws, f"{c}25", f"=RANK({c}23,$C$23:$F$23,1)", "calc", "int")
        W(ws, f"{c}26", f"=11-{c}23", "out", "sc")
    ws.conditional_formatting.add("C11:F21", ColorScaleRule(start_type="num", start_value=1, start_color="63BE7B", mid_type="num", mid_value=5.5, mid_color="FFEB84", end_type="num", end_value=10, end_color="F8696B"))
    ws.conditional_formatting.add("C24:F24", CellIsRule(operator="equal", formula=['"High"'], fill=PatternFill("solid", bgColor=RED), font=Font(color=DRED, bold=True)))
    ws.conditional_formatting.add("C24:F24", CellIsRule(operator="equal", formula=['"Low"'], fill=PatternFill("solid", bgColor=GREEN)))
    R["risk_w"] = {a: f"Risk_Analysis!${CL(3+i)}$23" for i, a in enumerate(APPS)}
    R["risk_m"] = {a: f"Risk_Analysis!${CL(3+i)}$26" for i, a in enumerate(APPS)}
    R["risk_wsum"] = "Risk_Analysis!$B$22"
    section(ws, 28, "2. Risk register for the recommended strategy (Industrial -> Robotics & Physical AI)", 11)
    header(ws, 29, ["Risk", "Probability", "Impact", "P score", "I score", "Exposure (PxI)", "Warning indicators", "Mitigation strategy", "Mitigation effectiveness", "Residual exposure", "Source"], h=34)
    dvp = DataValidation(type="list", formula1='"Low,Medium,High"', allow_blank=False); dvi = DataValidation(type="list", formula1='"Low,Medium,High,Severe"', allow_blank=False)
    ws.add_data_validation(dvp); ws.add_data_validation(dvi)
    for i, (rk, p, im, wn, mt, ef, src) in enumerate(RISK_REG):
        r = 30 + i
        W(ws, f"A{r}", rk, "label", wrap=True)
        W(ws, f"B{r}", p, "src" if "Report" in src else "assump", h="center"); W(ws, f"C{r}", im, "src" if "Report" in src else "assump", h="center")
        dvp.add(f"B{r}"); dvi.add(f"C{r}")
        W(ws, f"D{r}", f"=INDEX($B$41:$B$43,MATCH(B{r},$A$41:$A$43,0))", "calc", "int")
        W(ws, f"E{r}", f"=INDEX($E$41:$E$44,MATCH(C{r},$D$41:$D$44,0))", "calc", "int")
        W(ws, f"F{r}", f"=D{r}*E{r}", "calc", "int")
        W(ws, f"G{r}", wn, "grey", wrap=True); W(ws, f"H{r}", mt, "grey", wrap=True)
        W(ws, f"I{r}", ef, "assump", "pct")
        W(ws, f"J{r}", f"=F{r}*(1-I{r})", "calc", "0.0")
        W(ws, f"K{r}", src, "grey", wrap=True)
        ws.row_dimensions[r].height = 44
    W(ws, "A37", "Average exposure (gross / residual)", "label", bold=True)
    W(ws, "F37", "=AVERAGE(F30:F36)", "out", "0.0"); W(ws, "J37", "=AVERAGE(J30:J36)", "out", "0.0")
    W(ws, "A38", "Mitigation effectiveness is an ANALYST ASSUMPTION (share of exposure removed); the report supplies the mitigation actions but no effectiveness estimates.", "note", border=False)
    ws.conditional_formatting.add("F30:F36", ColorScaleRule(start_type="num", start_value=1, start_color="63BE7B", mid_type="num", mid_value=6, mid_color="FFEB84", end_type="num", end_value=12, end_color="F8696B"))
    section(ws, 39, "Lookup tables (probability and impact scales)", 11)
    header(ws, 40, ["Probability", "Score", "", "Impact", "Score"])
    for i, (lab, s) in enumerate((("Low", 1), ("Medium", 2), ("High", 3))):
        W(ws, f"A{41+i}", lab, "label"); W(ws, f"B{41+i}", s, "assump", "int")
    for i, (lab, s) in enumerate((("Low", 1), ("Medium", 2), ("High", 3), ("Severe", 4))):
        W(ws, f"D{41+i}", lab, "label"); W(ws, f"E{41+i}", s, "assump", "int")
    R["risk_reg_avg"], R["risk_reg_res"] = "Risk_Analysis!$F$37", "Risk_Analysis!$J$37"


# --------------------------------------------------------------------------------------------- FORECAST
FC_INPUT = {  # app: (prob b/bu/be, growth b/bu/be, GM change b/bu/be pp)
    "Automotive": ([.55, .15, .30], [.03, .07, -.08], [0.0, .015, -.03]),
    "Consumer": ([.50, .15, .35], [.02, .08, -.10], [0.0, .015, -.03]),
    "IoT": ([.55, .20, .25], [.06, .12, -.08], [0.0, .015, -.03]),
}


def build_forecast(ws):
    title(ws, "FORECAST & SCENARIO ANALYSIS - FY2027 (next year)",
          "Industrial scenarios = research-report estimates (365/410/295; 61.5%/63.5%/55.0%; 60/20/20). Other applications = ILLUSTRATIVE analyst scenarios (yellow). Not guaranteed results.", 7)
    widths(ws, [54, 14, 14, 14, 18, 40, 70])
    ws.freeze_panes = "B4"
    section(ws, 4, "Global inputs", 7)
    W(ws, "A5", "Gross-margin-to-operating-margin pass-through (share of GM change reaching operating margin)", "label"); W(ws, "B5", 0.8, "assump", "pct")
    W(ws, "A6", "Other / unallocated revenue FY2027 (held flat at FY2026E)", "label"); W(ws, "B6", "=Revenue_Analysis!$E$10", "calc", "bn")
    W(ws, "F5", "ANALYST ASSUMPTION: opex growth absorbs 20% of any gross-margin movement", "note", border=False)
    R["fc_other"] = "Forecast_Scenarios!$B$6"
    R["fc_sum_rev"], R["fc_sum"] = {}, {}
    blocks = {"Automotive": 9, "Industrial": 22, "Consumer": 35, "IoT": 48}
    R["fc_blocks"] = {}
    for app, s in blocks.items():
        ind = app == "Industrial"
        label = {"Automotive": "AUTOMOTIVE (illustrative scenarios)", "Industrial": "INDUSTRIAL (estimate; scenarios from research report s.17)",
                 "Consumer": "CONSUMER (illustrative scenarios)", "IoT": "IoT & INFRASTRUCTURE (illustrative scenarios)"}[app]
        section(ws, s, label, 7)
        header(ws, s + 1, ["Metric", "Base", "Bull", "Bear", "Prob-weighted / check", "Data type", "Notes"], h=24)
        L = lambda k: s + 2 + k
        names = ["Scenario probability", "FY2026E base revenue (JPY bn)", "Revenue growth vs FY2026E", "FY2027 revenue (JPY bn)",
                 "Gross margin change vs FY2025 proxy (pp)", "Gross margin", "Gross profit (JPY bn)", "Operating margin", "Operating profit (JPY bn)"]
        for k, n in enumerate(names):
            W(ws, f"A{L(k)}", n, "label")
        pr, gr, gm = FC_INPUT.get(app, (None, None, None))
        bcol = ["B", "C", "D"]
        for j, c in enumerate(bcol):
            if ind:
                W(ws, f"{c}{L(0)}", f"={R['K' + str(16 + j)]}", "src", "pct")
            else:
                W(ws, f"{c}{L(0)}", pr[j], "assump", "pct")
            W(ws, f"{c}{L(1)}", f"=Revenue_Analysis!$E${R['rev_rows'][app]}", "calc", "bn")
            if ind:
                W(ws, f"{c}{L(2)}", f"={c}{L(3)}/{c}{L(1)}-1", "calc", "pct")
                W(ws, f"{c}{L(3)}", f"={R['K' + str(10 + j)]}", "src", "bn")
                W(ws, f"{c}{L(5)}", f"={R['K' + str(13 + j)]}", "src", "pct")
                W(ws, f"{c}{L(4)}", f"={c}{L(5)}-{R['gm_proxy'][app]}", "calc", "pp")
            else:
                W(ws, f"{c}{L(2)}", gr[j], "assump", "pp")
                W(ws, f"{c}{L(3)}", f"={c}{L(1)}*(1+{c}{L(2)})", "calc", "bn")
                W(ws, f"{c}{L(4)}", gm[j], "assump", "pp")
                W(ws, f"{c}{L(5)}", f"={R['gm_proxy'][app]}+{c}{L(4)}", "calc", "pct")
            W(ws, f"{c}{L(6)}", f"={c}{L(3)}*{c}{L(5)}", "calc", "bn")
            W(ws, f"{c}{L(7)}", f"={R['om_proxy'][app]}+$B$5*{c}{L(4)}", "calc", "pct")
            W(ws, f"{c}{L(8)}", f"={c}{L(3)}*{c}{L(7)}", "calc", "bn")
        pcell = f"$B${L(0)}:$D${L(0)}"
        W(ws, f"E{L(0)}", f'=SUM(B{L(0)}:D{L(0)})', "out", "pct")
        W(ws, f"E{L(2)}", f"=E{L(3)}/B{L(1)}-1", "out", "pct")
        W(ws, f"E{L(3)}", f"=SUMPRODUCT({pcell},B{L(3)}:D{L(3)})", "out", "bn")
        W(ws, f"E{L(5)}", f"=E{L(6)}/E{L(3)}", "out", "pct")
        W(ws, f"E{L(6)}", f"=SUMPRODUCT({pcell},B{L(6)}:D{L(6)})", "out", "bn")
        W(ws, f"E{L(7)}", f"=E{L(8)}/E{L(3)}", "out", "pct")
        W(ws, f"E{L(8)}", f"=SUMPRODUCT({pcell},B{L(8)}:D{L(8)})", "out", "bn")
        W(ws, f"E{L(4)}", "", "grey")
        W(ws, f"E{L(1)}", "", "grey")
        dt = {0: "Report (Industrial) / ANALYST (others)", 1: "Link", 2: "Derived (Ind) / ANALYST (others)", 3: "Report forecast (Ind) / illustrative", 4: "Derived / ANALYST",
              5: "Report proxy (Ind) / proxy", 6: "Calculated", 7: "PROXY (OM)", 8: "Calculated"}
        for k, t in dt.items():
            W(ws, f"F{L(k)}", t, "grey")
        W(ws, f"G{L(0)}", "Probabilities must sum to 100% (check in column E; conditional format turns red otherwise)", "note", border=False)
        W(ws, f"G{L(7)}", "Operating margin = FY2025 OM proxy + pass-through x change in gross margin. Proxy only - application OM not disclosed.", "note", border=False)
        ws.conditional_formatting.add(f"E{L(0)}", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
        R["fc_sum_rev"][app] = f"Forecast_Scenarios!$E${L(3)}"
        R["fc_blocks"][app] = dict(prob=L(0), base=L(1), growth=L(2), rev=L(3), gmchg=L(4), gm=L(5), gp=L(6), om=L(7), op=L(8))
    # summary
    S = 61
    section(ws, S, "SUMMARY - probability-weighted FY2027 outcomes by application", 7)
    header(ws, S + 1, ["Application", "FY2026E revenue", "FY2027E expected revenue", "Expected growth", "Expected gross profit", "Expected GM", "Expected op. profit / OM"], h=34)
    for i, app in enumerate(APPS):
        r = S + 2 + i; b = R["fc_blocks"][app]
        W(ws, f"A{r}", app, "label", bold=True)
        W(ws, f"B{r}", f"=B{b['base']}", "calc", "bn")
        W(ws, f"C{r}", f"=E{b['rev']}", "out", "bn")
        W(ws, f"D{r}", f"=E{b['growth']}", "out", "pct")
        W(ws, f"E{r}", f"=E{b['gp']}", "calc", "bn")
        W(ws, f"F{r}", f"=E{b['gm']}", "calc", "pct")
        W(ws, f"G{r}", f"=E{b['op']}", "calc", "bn")
    W(ws, f"A{S+6}", "Total (four applications + Other flat)", "label", bold=True)
    W(ws, f"B{S+6}", f"=SUM(B{S+2}:B{S+5})+B6", "calc", "bn"); W(ws, f"C{S+6}", f"=SUM(C{S+2}:C{S+5})+B6", "out", "bn")
    W(ws, f"D{S+6}", f"=C{S+6}/B{S+6}-1", "out", "pct"); W(ws, f"E{S+6}", f"=SUM(E{S+2}:E{S+5})", "calc", "bn")
    W(ws, f"G{S+6}", f"=SUM(G{S+2}:G{S+5})", "calc", "bn")
    W(ws, f"A{S+7}", "Probability checks (each application sums to 100%)", "label")
    W(ws, f"B{S+7}", "=IF(AND(" + ",".join(f"ROUND(E{R['fc_blocks'][a]['prob']},6)=1" for a in APPS) + '),"PASS","FAIL")', "out", h="center")
    R["fc_prob_check"] = f"Forecast_Scenarios!$B${S+7}"
    R["fc_sum_growth"] = {a: f"Forecast_Scenarios!$D${S+2+i}" for i, a in enumerate(APPS)}
    R["fc_sum_gp"] = {a: f"Forecast_Scenarios!$E${S+2+i}" for i, a in enumerate(APPS)}
    R["fc_sum_gm"] = {a: f"Forecast_Scenarios!$F${S+2+i}" for i, a in enumerate(APPS)}
    R["fc_sum_op"] = {a: f"Forecast_Scenarios!$G${S+2+i}" for i, a in enumerate(APPS)}
    bi = R["fc_blocks"]["Industrial"]
    R["ind_rev"] = {k: f"Forecast_Scenarios!${c}${bi['rev']}" for k, c in zip(("base", "bull", "bear"), "BCD")}
    R["ind_gm"] = {k: f"Forecast_Scenarios!${c}${bi['gm']}" for k, c in zip(("base", "bull", "bear"), "BCD")}
    R["ind_prob"] = {k: f"Forecast_Scenarios!${c}${bi['prob']}" for k, c in zip(("base", "bull", "bear"), "BCD")}
    R["ind_exp_rev"] = f"Forecast_Scenarios!$E${bi['rev']}"
    R["ind_exp_gp"] = f"Forecast_Scenarios!$E${bi['gp']}"
    R["ind_base26"] = f"Forecast_Scenarios!$B${bi['base']}"
    R["ind_exp_growth"] = f"Forecast_Scenarios!$E${bi['growth']}"
    section(ws, S + 9, "What would cause each scenario to materialise", 7)
    header(ws, S + 10, ["Scenario", "Industrial trigger (research report s.17)", "", "", "", "Other applications (analyst, illustrative)", ""], h=24)
    trig = [("Base", "Steady industrial recovery; Altium integration lifts 'Winning Combination' design-wins ~15% among mid-tier robotics OEMs; GaN penetration in robotics motors proceeds smoothly.",
             "Auto: flat-to-modest volume with price-downs offset by content; Consumer: soft demand; IoT/Infra: steady connectivity & data-centre pull."),
            ("Bull", "Rapid proliferation of physical AI in warehousing/logistics; DRP-AI pull-through; Altium becomes de-facto AI-hardware design standard; GM expands to 63.5%.",
             "Auto: faster EV/ADAS content ramp; Consumer: restocking; IoT: AI data-centre power and edge-AI demand."),
            ("Bear", "Global industrial capex freeze + geopolitics; Chinese vendors qualify alternative MCUs, price war in mid-tier robotics; GM falls to 55.0%.",
             "Auto: renewed inventory correction + Chinese price war; Consumer: demand slump; IoT: channel destocking.")]
    for i, (sc, t1, t2) in enumerate(trig):
        r = S + 11 + i
        W(ws, f"A{r}", sc, "label", bold=True); W(ws, f"B{r}", t1, "grey", wrap=True); W(ws, f"F{r}", t2, "grey", wrap=True)
        ws.merge_cells(f"B{r}:E{r}"); ws.merge_cells(f"F{r}:G{r}")
        ws.row_dimensions[r].height = 48
    W(ws, f"A{S+15}", "Note: report ties the 365 bn base case to the whole pure-Industrial estimate; it gives NO revenue split for Robotics & Physical AI and no baseline. Revenue attributable to the incremental robotics programme is modelled separately in Resource_Allocation (assumptions, not report figures).", "note", border=False)


# --------------------------------------------------------------------------------------------- GROWTH
OPP_FACTORS = [("Market growth (company-relevant end market)", .20, [4, 8, 3, 7]),
               ("Semiconductor content growth per unit", .20, [8, 8, 3, 6]),
               ("Technology adoption (edge AI, electrification, connectivity)", .20, [6, 8, 4, 7]),
               ("Customer demand visibility", .20, [5, 7, 4, 6]),
               ("Addressable market size", .20, [8, 8, 6, 7])]


def build_growth(ws):
    title(ws, "GROWTH & MARKET OPPORTUNITY", "Market CAGR (third-party, via report) is NOT Renesas revenue CAGR. A growing market does not automatically translate into company revenue.", 9)
    widths(ws, [46, 12, 14, 12, 14, 14, 16, 18, 70])
    ws.freeze_panes = "B4"
    section(ws, 4, "1. Market growth evidence from the report (TOTAL markets - not Renesas-addressable semiconductor content)", 9)
    header(ws, 5, ["Market", "Start year", "Start size (USD bn)", "End year", "End size (USD bn)", "Years", "CAGR (derived)", "CAGR stated by report", "Comment"], h=34)
    mk = [("Factory automation & industrial control", 2025, "M01", 2030, "M02", None, "Derived CAGR ~6.9% is far below robotics headline growth; large but slower market."),
          ("AI infrastructure / data centres", 2024, "M03", 2030, "M04", None, "Fast and price-inelastic, but concentrated hyperscaler customers and power-only pull-through (s.6)."),
          ("Physical AI / robotics (total market)", None, None, 2035, "M05", "M06", "NO base-year size is given, so a CAGR cannot be derived; report quotes '>30%' (s.6) and '>25%' (s.11) - inconsistent (F-07)."),
          ("Analog IC", 2025, "M08", 2035, "M10", "M09", "Derived CAGR reproduces the stated 3.0%: mature market (consistency check)."),
          ("GaN power semiconductors", None, None, None, None, "M12", "Stated only; no market size given.")]
    for i, (n, y0, s0, y1, s1, st, cm) in enumerate(mk):
        r = 6 + i
        W(ws, f"A{r}", n, "label")
        W(ws, f"B{r}", y0 if y0 else "n/a", "src" if y0 else "grey", "0", h="center")
        W(ws, f"C{r}", f"={R[s0]}" if s0 else "n/a", "calc" if s0 else "grey", "bn", h="center")
        W(ws, f"D{r}", y1 if y1 else "n/a", "src" if y1 else "grey", "0", h="center")
        W(ws, f"E{r}", f"={R[s1]}" if s1 else "n/a", "calc" if s1 else "grey", "bn", h="center")
        if y0 and y1:
            W(ws, f"F{r}", f"=D{r}-B{r}", "calc", "int")
            W(ws, f"G{r}", f"=(E{r}/C{r})^(1/F{r})-1", "out", "pct")
        else:
            W(ws, f"F{r}", "n/a", "grey", h="center"); W(ws, f"G{r}", "n/a", "grey", h="center")
        W(ws, f"H{r}", f"={R[st]}" if st else "n/a", "calc" if st else "grey", "pct", h="center")
        W(ws, f"I{r}", cm, "grey", wrap=True)
        ws.row_dimensions[r].height = 30
    R["g_fa_cagr"], R["g_ai_cagr"] = "Growth_Opportunity!$G$6", "Growth_Opportunity!$G$7"
    section(ws, 12, "2. Company revenue growth vs market growth (kept separate)", 9)
    header(ws, 13, ["Application", "Hist. company CAGR FY23-25", "FY26E YoY (est.)", "FY27E expected growth (scenario)", "", "", "", "", "Reading"], h=44)
    reads = {"Automotive": "Contracting FY24-25 then rebound (est.); expected growth low in scenarios.",
             "Industrial": "Hist. growth = I&I by construction; forward growth comes from the report's scenarios (pro-Industrial source) - see bias note.",
             "Consumer": "Illustrative; identical history to I&I by construction.", "IoT": "Illustrative; includes infrastructure/data-centre exposure."}
    for i, app in enumerate(APPS):
        r = 14 + i
        W(ws, f"A{r}", app, "label", bold=True)
        W(ws, f"B{r}", f"=Revenue_Analysis!$F${R['rev_rows'][app]}", "calc", "pct")
        W(ws, f"C{r}", f"=Revenue_Analysis!$E${R['rev_rows'][app]}/Revenue_Analysis!$D${R['rev_rows'][app]}-1", "calc", "pct")
        W(ws, f"D{r}", f"={R['fc_sum_growth'][app]}", "calc", "pct")
        W(ws, f"I{r}", reads[app], "grey", wrap=True); ws.row_dimensions[r].height = 28
    W(ws, "A18", "Bias note: Industrial forward growth is taken from the pro-Industrial research report; Automotive/Consumer/IoT scenarios are analyst-set. The recommendation is therefore stress-tested in Sensitivity (haircut on Industrial).", "note", border=False)
    section(ws, 20, "3. Market-opportunity factor scores (1-10, higher = more attractive) - ANALYST ASSESSMENT", 9)
    header(ws, 21, ["Factor", "Weight"] + APPS + ["", "", "Basis"], h=28)
    basis = ["Market growth: auto EV deceleration (s.1); Industrial robotics/edge-AI fast; consumer saturated; IoT/data-centre strong (s.6, s.11).",
             "Content: ADAS/electrification raises auto content; robots need MCU+MPU+PMIC+power (s.6).",
             "Adoption: edge AI inference, cobots, predictive maintenance (s.5, s.6).",
             "Demand visibility: long-lifecycle design-wins (s.3) vs consumer re-bids.",
             "TAM: factory automation USD 312-435bn; AI infrastructure USD 132-371bn; robotics smaller but fastest."]
    for i, (f, w, sc) in enumerate(OPP_FACTORS):
        r = 22 + i
        W(ws, f"A{r}", f, "label"); W(ws, f"B{r}", w, "assump", "pct")
        for j, s in enumerate(sc):
            W(ws, f"{CL(3+j)}{r}", s, "assump", "int")
        W(ws, f"I{r}", basis[i], "grey", wrap=True); ws.row_dimensions[r].height = 28
    W(ws, "A27", "Weights check (must be 100%)", "label"); W(ws, "B27", "=SUM(B22:B26)", "calc", "pct")
    W(ws, "A28", "MARKET-OPPORTUNITY SCORE", "label", bold=True)
    W(ws, "A29", "Rank", "label")
    for j in range(4):
        c = CL(3 + j)
        W(ws, f"{c}28", f"=SUMPRODUCT($B$22:$B$26,{c}22:{c}26)", "out", "sc")
        W(ws, f"{c}29", f"=RANK({c}28,$C$28:$F$28)", "calc", "int")
    ws.conditional_formatting.add("C22:F26", ColorScaleRule(start_type="num", start_value=1, start_color="F8696B", mid_type="num", mid_value=5.5, mid_color="FFEB84", end_type="num", end_value=10, end_color="63BE7B"))
    R["opp"] = {a: f"Growth_Opportunity!${CL(3+i)}$28" for i, a in enumerate(APPS)}
    R["opp_wsum"] = "Growth_Opportunity!$B$27"
    section(ws, 31, "4. Renesas-specific ability to capture the opportunity (competitive position)", 9)
    header(ws, 32, ["Application", "Principal competitors", "", "", "", "", "", "", "Renesas position / evidence"], h=24)
    cap = [("Automotive", "NXP, Infineon, STMicro, TI; Chinese entrants", "Historic stronghold (~48% of FY2025 revenue); but SiC/IGBT R&D suspended after Wolfspeed RSA, limiting EV-powertrain content (s.1, s.10, s.18)."),
           ("Industrial", "TI, ADI, Infineon, STMicro", "RZ MPUs, RA8 MCUs, Dialog PMICs, Transphorm GaN, Reality AI, Altium: strongest integrated stack; but TI/Infineon entrenched in factory automation (s.5, s.10)."),
           ("Consumer", "Chinese MCU/PMIC suppliers (SG Micro, GigaDevice), MediaTek, STM", "Commoditised; price-led; report excludes it from Industrial resources (s.4, s.18)."),
           ("IoT", "STM32 ecosystem, Nordic, Espressif, TI", "Celeno Wi-Fi 6/6E + BLE + RA MCUs competitive but low content per device; building automation commoditising (s.6).")]
    for i, (a, comp, ev) in enumerate(cap):
        r = 33 + i
        W(ws, f"A{r}", a, "label", bold=True); W(ws, f"B{r}", comp, "grey", wrap=True); W(ws, f"I{r}", ev, "grey", wrap=True)
        ws.merge_cells(f"B{r}:H{r}"); ws.row_dimensions[r].height = 40
    section(ws, 38, "5. Robotics & Physical AI - how the Renesas portfolio combines commercially", 9)
    header(ws, 39, ["Product / asset", "Role in robot", "", "", "", "", "", "Revenue-per-customer effect", "Evidence and caveat (report section)"], h=28)
    eco = [("RA8 (Arm Cortex-M85, Helium)", "Deterministic motor control + edge inference", "Anchor design-win; sets the attach point", "s.5: inference without separate NPU. Caveat: RISC-V substitution risk (s.18)."),
           ("RZ/V MPUs with DRP-AI", "Machine vision / spatial mapping", "Higher-ASP processor in vision robots", "s.5/s.6: inference-per-watt claim is vendor-sourced [8]; needs independent benchmarking."),
           ("Dialog PMICs", "Dense power sequencing for MCU/MPU rails", "Attach #1: PMIC per processing node", "s.5/s.8: bundling blocks modular substitution; Chinese PMIC undercut 30-40% (s.18)."),
           ("Transphorm GaN (d-mode)", "48V joint / motor-drive power stage", "Attach #2: power stage per joint", "s.5: 650V+ GaN; robot-joint volume unquantified; AI data centre is main GaN pull (s.6)."),
           ("Reality AI software", "Train/deploy ML models on Renesas MCUs", "Reduces customer design friction; locks in software layer", "s.5: revenue not disclosed; small versus hardware."),
           ("Altium (PCB design platform)", "Design-entry funnel for 500 target developers", "Pipeline source; recurring licence revenue", "s.3: USD 263m revenue, 77% recurring (~3% of group). Vendor-neutral tool - 'workflow monopoly' claim is overstated (F-09)."),
           ("Celeno Wi-Fi / BLE", "Connectivity for mobile robots", "Optional attach", "s.5: supporting role."),
           ("Winning Combinations", "Pre-validated multi-chip reference designs", "Raises content per design-win & switching cost", "s.3/s.8: >35 combinations; attach target >2.5 ICs per processor (K20).")]
    for i, (p, role, eff, ev) in enumerate(eco):
        r = 40 + i
        W(ws, f"A{r}", p, "label", bold=True); W(ws, f"B{r}", role, "grey", wrap=True); W(ws, f"H{r}", eff, "grey", wrap=True); W(ws, f"I{r}", ev, "grey", wrap=True)
        ws.merge_cells(f"B{r}:G{r}"); ws.row_dimensions[r].height = 40
    section(ws, 49, "6. ILLUSTRATIVE attach-rate content model (INDEX units - no ASPs are disclosed; this is not a revenue forecast)", 9)
    W(ws, "A50", "Baseline attach rate (analog/power ICs per processing node) - ASSUMPTION", "label"); W(ws, "B50", 1.0, "assump", "0.0")
    W(ws, "A51", "Target attach rate (report KPI, s.21)", "label"); W(ws, "B51", f"={R['K20']}", "src", "0.0")
    W(ws, "A52", "Relative ASP of an attached IC vs the processor (processor = 1.00) - ASSUMPTION", "label"); W(ws, "B52", 0.30, "assump", "0.00")
    W(ws, "A53", "Content per node at baseline (index)", "label"); W(ws, "B53", "=1+B50*B52", "calc", "0.00")
    W(ws, "A54", "Content per node at target attach (index)", "label"); W(ws, "B54", "=1+B51*B52", "calc", "0.00")
    W(ws, "A55", "Uplift in semiconductor content per design-win", "label", bold=True); W(ws, "B55", "=B54/B53-1", "out", "pct")
    R["attach_uplift"] = "Growth_Opportunity!$B$55"
    W(ws, "C55", "Same design-win, more Renesas content: the commercial logic of Winning Combinations. The size depends on the assumed relative ASP (grid below).", "note", border=False)
    header(ws, 57, ["Uplift: target attach (down) x relative ASP (across)", "0.20", "0.30", "0.40", "0.50"], h=28)
    for j, v in enumerate((0.20, 0.30, 0.40, 0.50)):
        ws.cell(row=57, column=2 + j).value = v
        ws.cell(row=57, column=2 + j).number_format = "0.00"
    for i, t in enumerate((1.5, 2.0, 2.5, 3.0)):
        r = 58 + i
        W(ws, f"A{r}", t, "assump", "0.0", h="center")
        for j in range(4):
            c = CL(2 + j)
            W(ws, f"{c}{r}", f"=(1+$A{r}*{c}$57)/(1+$B$50*{c}$57)-1", "calc", "pct")
    ws.conditional_formatting.add("B58:E61", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="63BE7B"))
    section(ws, 63, "7. Alternatives within Industrial (scores from Strategic_Scorecard Part B)", 9)
    text_rows(ws, 64, ["See Strategic_Scorecard Part B: Robotics & Physical AI ranks first, AI data centres second, with the margin depending on the report's own subjective Competitive-advantage and Strategic-alignment scores (see Sensitivity)."])

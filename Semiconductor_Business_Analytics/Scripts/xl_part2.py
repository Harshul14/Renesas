"""Workbook sheets part 2: Strategic_Scorecard, Resource_Allocation, Sensitivity, Data_Validation, Sources_Assumptions, Dashboard."""
from openpyxl.chart import BarChart, ScatterChart, DoughnutChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment
from common import *
from xl_part1 import R, SUBMARKETS, SUB_CRIT, rawy

CRIT = [  # name, weight, scores or None (linked), type, basis
    ("Revenue growth potential", .20, [4, 8, 3, 7], "Analyst",
     "WEIGHT: growth is the stated objective, so the largest weight. SCORES: forward scenario growth (Forecast_Scenarios), history and market growth. Industrial 8 (report-based, probability-weighted ~+12%); Automotive 4; Consumer 3; IoT 7."),
    ("Profitability potential", .15, [6, 7, 4, 6], "Analyst (proxy-based)",
     "WEIGHT: profit converts growth into value. SCORES use Profitability_Quality proxies: Industrial highest gross margin proxy; Automotive highest implied operating margin (derived); Consumer lowest."),
    ("Market opportunity", .15, None, "LINKED: Growth_Opportunity s.3", "WEIGHT: sizing of the prize. SCORE = weighted market-opportunity score from Growth_Opportunity (factor scores are analyst assessments)."),
    ("Competitive advantage", .15, [7, 7, 4, 6], "Analyst",
     "WEIGHT: ability to win, not just market size. SCORES: Industrial 7 (integrated stack but TI/Infineon entrenched); Automotive 7 (stronghold but SiC retreat); IoT 6; Consumer 4 (commoditised)."),
    ("Revenue quality", .10, None, "LINKED: Profitability_Quality s.5", "WEIGHT: predictability matters but is partly captured via risk. SCORE = weighted revenue-quality score (stability, concentration, recurring, retention, lifecycle, pricing power, switching cost, capital intensity)."),
    ("Speed of revenue conversion", .10, [4, 5, 7, 7], "Analyst",
     "WEIGHT: next-year focus. SCORES: Consumer/IoT shorter cycles (7); Industrial 12-24 months, up to 36 for factory automation (5); Automotive multi-year qualification (4, analyst)."),
    ("Strategic alignment", .10, [5, 9, 3, 6], "Analyst",
     "WEIGHT: fit with platform strategy (Winning Combinations, Altium, GaN pivot). SCORES: Industrial 9; IoT 6; Automotive 5 (SiC/IGBT R&D suspended); Consumer 3. Subjective - stress-tested in Sensitivity."),
    ("Risk manageability (high = LOWER risk)", .05, None, "LINKED: 11 - Risk_Analysis weighted risk", "WEIGHT: low at 5% as suggested, because risk is also embedded in quality/competition/speed; Sensitivity raises it to 40%. SCORE = 11 - weighted risk score."),
]
SC_FIRST, SC_LAST, SC_TOT = 6, 13, 26
SUBC = ["Revenue growth", "Addressable market", "Competitive advantage", "Profitability", "Speed of conversion", "Design-win visibility",
        "Strategic alignment", "Risk manageability", "Investment feasibility"]


def build_scorecard(ws):
    title(ws, "WEIGHTED STRATEGIC SCORECARD", "All scores 1-10, HIGHER = MORE ATTRACTIVE. Weights are editable yellow inputs (must total 100%). Part A: four applications. Part B: Industrial submarkets.", 12)
    widths(ws, [44, 12, 15, 14, 14, 14, 14, 14, 14, 14, 14, 14])
    ws.freeze_panes = "B4"
    section(ws, 4, "PART A - Four applications (Automotive | Industrial | Consumer | IoT): criterion scores and weights", 12)
    header(ws, 5, ["Criterion", "Weight"] + APPS + ["Score type", "Basis / justification of weight and scores", "", "", "", ""], h=30)
    for i, (n, w, sc, tp, bs) in enumerate(CRIT):
        r = SC_FIRST + i
        W(ws, f"A{r}", n, "label"); W(ws, f"B{r}", w, "assump", "pct")
        for j, app in enumerate(APPS):
            c = CL(3 + j)
            if sc is not None:
                W(ws, f"{c}{r}", sc[j], "assump", "int")
            else:
                f = {2: R["opp"][app], 4: R["quality"][app], 7: R["risk_m"][app]}[i]
                W(ws, f"{c}{r}", f"={f}", "calc", "sc")
        W(ws, f"G{r}", tp, "grey", wrap=True)
        W(ws, f"H{r}", bs, "grey", wrap=True)
        ws.merge_cells(f"H{r}:L{r}")
        ws.row_dimensions[r].height = 56
    W(ws, "A14", "Weights check (must be 100%)", "label"); W(ws, "B14", "=SUM(B6:B13)", "out", "pct")
    ws.conditional_formatting.add("B14", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
    ws.conditional_formatting.add("C6:F13", ColorScaleRule(start_type="num", start_value=1, start_color="F8696B", mid_type="num", mid_value=5.5, mid_color="FFEB84", end_type="num", end_value=10, end_color="63BE7B"))
    dv = DataValidation(type="decimal", operator="between", formula1="1", formula2="10", showErrorMessage=True, errorTitle="Score out of range", error="Scores must be between 1 and 10")
    ws.add_data_validation(dv)
    for i, (n, w, sc, tp, bs) in enumerate(CRIT):
        if sc is not None:
            dv.add(f"C{SC_FIRST+i}:F{SC_FIRST+i}")
    dvw = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True, error="Weights are percentages between 0% and 100%")
    ws.add_data_validation(dvw); dvw.add("B6:B13")
    W(ws, "A15", "Suggested weights follow the assignment brief; they are proposals, not facts. Linked rows (green font in grey) update automatically when their source sheets change.", "note", border=False)
    section(ws, 16, "Weighted contributions = score x weight", 12)
    header(ws, 17, ["Criterion", "Weight"] + APPS)
    for i in range(8):
        r = 18 + i
        W(ws, f"A{r}", f"=A{SC_FIRST+i}", "label"); W(ws, f"B{r}", f"=B{SC_FIRST+i}", "calc", "pct")
        for j in range(4):
            c = CL(3 + j)
            W(ws, f"{c}{r}", f"={c}{SC_FIRST+i}*$B{SC_FIRST+i}", "calc", "sc")
    W(ws, f"A{SC_TOT}", "WEIGHTED SCORE = SUMPRODUCT(scores, weights)", "label", bold=True)
    W(ws, "A27", "Rank (1 = most attractive)", "label")
    W(ws, "A28", "Gap to rank-1 application (points)", "label")
    W(ws, "A29", "Strategic interpretation", "label")
    W(ws, "A30", "Top-ranked application", "label", bold=True)
    W(ws, "A31", "Lead of top application over runner-up (points)", "label")
    W(ws, "A32", "Cross-check: sum of contributions less SUMPRODUCT (must be 0.00)", "label")
    for j in range(4):
        c = CL(3 + j)
        W(ws, f"{c}{SC_TOT}", f"=SUMPRODUCT($B${SC_FIRST}:$B${SC_LAST},{c}{SC_FIRST}:{c}{SC_LAST})", "out", "sc")
        W(ws, f"{c}27", f"=RANK({c}{SC_TOT},$C${SC_TOT}:$F${SC_TOT})", "out", "int")
        W(ws, f"{c}28", f"=MAX($C${SC_TOT}:$F${SC_TOT})-{c}{SC_TOT}", "calc", "sc")
        W(ws, f"{c}29", f'=IF({c}27=1,"PRIORITY: majority of incremental resources",IF({c}27=2,"SELECTIVE: targeted programmes",IF({c}27=3,"MAINTAIN / DEFEND","DE-PRIORITISE")))', "calc", wrap=True)
        W(ws, f"{c}32", f"=ROUND(SUM({c}18:{c}25)-{c}{SC_TOT},9)", "calc", "sc")
    ws.row_dimensions[29].height = 44
    W(ws, "C30", f"=INDEX($C$5:$F$5,MATCH(MAX($C${SC_TOT}:$F${SC_TOT}),$C${SC_TOT}:$F${SC_TOT},0))", "out", h="center")
    W(ws, "C31", f"=LARGE($C${SC_TOT}:$F${SC_TOT},1)-LARGE($C${SC_TOT}:$F${SC_TOT},2)", "out", "sc")
    ws.conditional_formatting.add(f"C{SC_TOT}:F{SC_TOT}", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    R["sc_total"] = {a: f"Strategic_Scorecard!${CL(3+i)}${SC_TOT}" for i, a in enumerate(APPS)}
    R["sc_rank"] = {a: f"Strategic_Scorecard!${CL(3+i)}$27" for i, a in enumerate(APPS)}
    R["sc_winner"], R["sc_lead"] = "Strategic_Scorecard!$C$30", "Strategic_Scorecard!$C$31"
    R["sc_wsum"] = "Strategic_Scorecard!$B$14"
    R["sc_interp"] = {a: f"Strategic_Scorecard!${CL(3+i)}$29" for i, a in enumerate(APPS)}

    # ---------------- Part B
    section(ws, 34, "PART B - Industrial submarkets: research-report scores recomputed (SUMPRODUCT) under three weight sets", 12)
    W(ws, "A35", "Weight set used for headline ranking (1 = report weights, 2 = assignment weights mapped, 3 = custom)", "label", wrap=True)
    ws.row_dimensions[35].height = 30
    W(ws, "B35", 2, "assump", "int", h="center")
    dvs = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False); ws.add_data_validation(dvs); dvs.add("B35")
    W(ws, "C35", '=CHOOSE(B35,"Report weights (as published)","Assignment weights mapped: quality<-design-win visibility; investment feasibility 0%","Custom weights (column D)")', "calc", border=False)
    header(ws, 37, ["Criterion", "Report weights", "Assignment (mapped)", "Custom", "SELECTED"] + SUBMARKETS, h=48)
    amap = [.20, .15, .15, .15, .10, .10, .10, .05, 0.0]
    r0 = R["sub_first_row"]
    for i in range(9):
        r = 38 + i
        W(ws, f"A{r}", SUBC[i] if i != 7 else "Risk manageability (high = lower risk)", "label")
        W(ws, f"B{r}", f"=Raw_Data!$C${r0+i}", "src", "pct")
        W(ws, f"C{r}", amap[i], "assump", "pct")
        W(ws, f"D{r}", SUB_CRIT[i][1], "assump", "pct")
        W(ws, f"E{r}", f"=CHOOSE($B$35,B{r},C{r},D{r})", "calc", "pct")
        for j in range(7):
            W(ws, f"{CL(6+j)}{r}", f"=Raw_Data!${CL(4+j)}${r0+i}", "src", "int")
    W(ws, "A47", "Weights check (must be 100%)", "label")
    for c in "BCDE":
        W(ws, f"{c}47", f"=SUM({c}38:{c}46)", "out", "pct")
    labs = [(48, "WEIGHTED SCORE - SELECTED weights", "E"), (49, "Rank - selected", None), (50, "Score - report weights", "B"), (51, "Score - assignment weights", "C"),
            (52, "Score - custom weights", "D"), (53, "Weighted total AS REPORTED in the report", None), (54, "Recomputed (report weights) less reported", None),
            (55, "Reconciliation flag (|diff| > 0.005)", None)]
    for r, t, wc in labs:
        W(ws, f"A{r}", t, "label", bold=(r in (48,)))
    for j in range(7):
        c = CL(6 + j)
        for r, wc in ((48, "E"), (50, "B"), (51, "C"), (52, "D")):
            W(ws, f"{c}{r}", f"=SUMPRODUCT(${wc}$38:${wc}$46,{c}38:{c}46)", "out" if r == 48 else "calc", "sc")
        W(ws, f"{c}49", f"=RANK({c}48,$F$48:$L$48)", "out", "int")
        W(ws, f"{c}53", f"=Raw_Data!${CL(4+j)}${R['sub_total_row']}", "src", "0.00")
        W(ws, f"{c}54", f"={c}50-{c}53", "calc", "+0.00;[Red]-0.00;0.00")
        W(ws, f"{c}55", f'=IF(ABS({c}54)>0.005,"FLAG","OK")', "calc", h="center")
    ws.conditional_formatting.add("F55:L55", CellIsRule(operator="equal", formula=['"FLAG"'], fill=PatternFill("solid", bgColor=RED), font=Font(color=DRED, bold=True)))
    ws.conditional_formatting.add("F48:L48", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    for k, (lab, row) in enumerate((("Top submarket - report weights", 50), ("Top submarket - assignment weights", 51), ("Top submarket - custom weights", 52), ("Top submarket - SELECTED weights", 48))):
        r = 57 + k
        W(ws, f"A{r}", lab, "label", bold=(k == 3))
        W(ws, f"B{r}", f"=INDEX($F$37:$L$37,MATCH(MAX($F${row}:$L${row}),$F${row}:$L${row},0))", "out" if k == 3 else "calc", h="left")
    W(ws, "A61", "Lead of #1 over #2 submarket - SELECTED weights (points)", "label"); W(ws, "B61", "=LARGE($F$48:$L$48,1)-LARGE($F$48:$L$48,2)", "out", "sc")
    W(ws, "A62", "Reconciliation flags raised on report-stated totals (count)", "label"); W(ws, "B62", '=COUNTIF(F55:L55,"FLAG")', "out", "int")
    text_rows(ws, 63, ["Interpretation: Robotics & Physical AI ranks first under every weight set, AI data centres second. The report's own published totals contain arithmetic discrepancies (flagged); the ranking is unchanged after recomputation,",
                       "but the lead depends heavily on the report's subjective Competitive-advantage (9) and Strategic-alignment (10) scores for Robotics - these partly reflect the Altium/Reality AI thesis that is still unproven (see Sensitivity)."])
    R["sub_sel_row"] = 48
    R["sub_names_range"] = "Strategic_Scorecard!$F$37:$L$37"
    R["sub_sel_range"] = "Strategic_Scorecard!$F$48:$L$48"
    R["sub_winner"] = "Strategic_Scorecard!$B$60"
    R["sub_lead"] = "Strategic_Scorecard!$B$61"
    R["sub_flags"] = "Strategic_Scorecard!$B$62"
    R["sub_wsums"] = ["Strategic_Scorecard!$B$47", "Strategic_Scorecard!$C$47", "Strategic_Scorecard!$D$47", "Strategic_Scorecard!$E$47"]
    R["sub_score"] = {n: f"Strategic_Scorecard!${CL(6+j)}$48" for j, n in enumerate(SUBMARKETS)}
    R["sub_rank"] = {n: f"Strategic_Scorecard!${CL(6+j)}$49" for j, n in enumerate(SUBMARKETS)}
    R["sub_report_w"] = [f"Strategic_Scorecard!$B${38+i}" for i in range(9)]
    # chart: submarket ranking
    ch = BarChart(); ch.type = "bar"; ch.title = "Industrial submarket weighted score (selected weights)"
    ch.add_data(Reference(ws, min_col=6, max_col=12, min_row=48), from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=6, max_col=12, min_row=37))
    ch.legend = None; ch.height, ch.width = 8, 16
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
    ws.add_chart(ch, "H5") if False else None
    ws.add_chart(ch, "A66")


# --------------------------------------------------------------------------------------------- RESOURCES
CATS = [("Field Application Engineering (FAE)", .35, .30,
         "Robotics designs need hands-on motor-control, ML-deployment and thermal support; FAEs are the conversion engine for 12-18 month design-wins. Trimmed to 30% because FAE intensity is the main cost-to-win risk (report Q11) and Altium/Reality AI self-serve funnel should absorb early-stage support."),
        ("Ecosystem & Software Development", .25, .25,
         "Integrates Winning Combinations into Altium, ROS2/PyTorch support, Reality AI: raises attach rate (content per design-win) and architecture-agnostic stickiness. Highest cost-efficiency and risk reduction (RISC-V hedge)."),
        ("Strategic Account Development", .20, .20,
         "Task forces on top robotics OEMs / physical-AI startups, supply agreements and joint R&D: concentrates scarce resource on accounts with large volume potential; secures capacity commitments."),
        ("Technical Partnerships", .10, .10,
         "Sensor/LiDAR/vision partners and pre-validated modules shorten customer time-to-market; moderate impact, partner-dependent."),
        ("Product Marketing", .10, .05,
         "Inference-per-watt message is vendor-sourced and brand is already strong with target developers; most demand generation is delivered via Altium and reference designs, so cut from 10% to 5%."),
        ("Customer Acquisition & Retention", .00, .10,
         "NEW vs report: protects the ~285 bn Industrial base and early design-wins (renewals, channel visibility, DOI monitoring); reduces bear-case exposure to inventory corrections and price wars.")]
EVAL = [  # revenue impact, speed, risk reduction, cost efficiency
    [9, 7, 6, 4], [8, 5, 8, 9], [8, 6, 7, 5], [6, 6, 6, 7], [5, 6, 4, 7], [6, 7, 8, 6]]
APP_REC = {"Automotive": .25, "Industrial": .55, "Consumer": .08, "IoT": .12}
APP_WHY = {"Automotive": "DEFEND: ~48% of FY2025 revenue and the highest implied operating margin; protect with targeted account management rather than growth spend. Scorecard measures marginal attractiveness, not defend value.",
           "Industrial": "PRIORITY: top-ranked on the weighted scorecard; superior gross margin, longer lifecycles, fragmented customers and highest platform fit.",
           "Consumer": "MINIMAL: lowest score; commoditised, price-led, cyclical; serve through channel/distribution only.",
           "IoT": "SELECTIVE: second on score thanks to connectivity and infrastructure pull, but low content per device; fund through existing channels and cross-sell."}
SUB_REC = [.20, .12, .45, .02, .13, .04, .04]
SUB_WHY = ["AI data centres: fast GaN revenue but concentrated customers and no MCU/software pull-through -> opportunistic, sized to existing GaN channel.",
           "Factory automation: large but 36-month qualification and entrenched TI/Infineon -> targeted displacement only.",
           "Robotics & Physical AI: highest strategic fit; 12-18 month design-win; integrated Winning Combinations; fragmented customers.",
           "Energy & smart grids: SiC/IGBT retreat leaves little addressable content -> minimal.",
           "Industrial IoT: solid, fast conversion, uses existing RA MCU + Celeno stack -> secondary.",
           "Building automation: commoditising, low content per device -> minimal.",
           "Test & measurement: high margin but walled off by ADI/TI -> minimal."]


def build_resources(ws):
    title(ws, "RESOURCE ALLOCATION - incremental commercial resources for FY2027", "Report allocation (35/25/20/10/10) is evaluated, not assumed optimal. Recommended allocation is an ANALYST JUDGEMENT cross-checked by an evidence-implied allocation.", 8)
    widths(ws, [44, 14, 14, 14, 16, 16, 22, 90])
    ws.freeze_panes = "B4"
    section(ws, 4, "A. Allocation by commercial-resource category (must total 100%)", 8)
    header(ws, 5, ["Resource category", "Report allocation", "RECOMMENDED allocation", "Change vs report", "Evidence-implied %", "Budget (JPY bn, illustrative)", "Eval. score (section B)", "Why this share / revenue link / risk reduction / expected outcome"], h=44)
    for i, (n, rep, rec, why) in enumerate(CATS):
        r = 6 + i
        W(ws, f"A{r}", n, "label", wrap=True); W(ws, f"B{r}", rep, "src", "pct"); W(ws, f"C{r}", rec, "assump", "pct")
        W(ws, f"D{r}", f"=C{r}-B{r}", "calc", "pp")
        W(ws, f"E{r}", f"=POWER(G{r},$B$25)/SUMPRODUCT(POWER($G$6:$G$11,$B$25))", "calc", "pct")
        W(ws, f"F{r}", f"=C{r}*$B$51", "calc", "0.00")
        W(ws, f"G{r}", f"=F{18+i}", "calc", "sc")
        W(ws, f"H{r}", why, "grey", wrap=True)
        ws.row_dimensions[r].height = 58
    W(ws, "A12", "TOTAL (must be 100%)", "label", bold=True)
    for c in "BCE":
        W(ws, f"{c}12", f"=SUM({c}6:{c}11)", "out", "pct")
    W(ws, "D12", "=SUM(D6:D11)", "calc", "pp"); W(ws, "F12", "=SUM(F6:F11)", "out", "0.00")
    ws.conditional_formatting.add("B12:C12", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
    ws.conditional_formatting.add("E12", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
    R["alloc_rep_sum"], R["alloc_rec_sum"], R["alloc_imp_sum"] = "Resource_Allocation!$B$12", "Resource_Allocation!$C$12", "Resource_Allocation!$E$12"
    R["alloc_cats"] = [(n, f"Resource_Allocation!$C${6+i}", f"Resource_Allocation!$B${6+i}") for i, (n, _, _, _) in enumerate(CATS)]
    section(ws, 14, "B. Evaluation of each category (1-10; higher = better) - ANALYST ASSESSMENT", 8)
    header(ws, 15, ["Category", "Revenue impact", "Speed to revenue", "Risk reduction", "Cost efficiency / scalability", "Weighted score", "", "Comment"], h=34)
    W(ws, "A16", "Criterion weights (must be 100%)", "label")
    for j, w in enumerate((.35, .20, .25, .20)):
        W(ws, f"{CL(2+j)}16", w, "assump", "pct")
    W(ws, "F16", "=SUM(B16:E16)", "calc", "pct")
    W(ws, "A17", "(scores below)", "note", border=False)
    for i, (n, *_r) in enumerate(CATS):
        r = 18 + i
        W(ws, f"A{r}", n, "label", wrap=True)
        for j, s in enumerate(EVAL[i]):
            W(ws, f"{CL(2+j)}{r}", s, "assump", "int")
        W(ws, f"F{r}", f"=SUMPRODUCT($B$16:$E$16,B{r}:E{r})", "out", "sc")
        ws.row_dimensions[r].height = 28
    W(ws, "H18", "Revenue impact is highest for FAEs (direct conversion) but cost efficiency is lowest (headcount heavy).", "note", border=False)
    W(ws, "H19", "Ecosystem/software scales across thousands of developers at low marginal cost.", "note", border=False)
    W(ws, "A24", "Why the recommended % differs from the evidence-implied %: FAE capacity is a hard constraint for 12-18 month design-wins, and recommended shares are rounded management-feasible blocks.", "note", border=False)
    W(ws, "A25", "Sharpening exponent p (evidence-implied % = score^p / sum) - ANALYST ASSUMPTION", "label"); W(ws, "B25", 4, "assump", "int")
    ws.conditional_formatting.add("B18:E23", ColorScaleRule(start_type="num", start_value=1, start_color="F8696B", mid_type="num", mid_value=5.5, mid_color="FFEB84", end_type="num", end_value=10, end_color="63BE7B"))
    section(ws, 27, "C. Allocation across the four applications (share of incremental commercial resources)", 8)
    header(ws, 28, ["Application", "Scorecard score", "Score-implied %", "RECOMMENDED %", "Rank", "", "", "Rationale"], h=30)
    for i, app in enumerate(APPS):
        r = 29 + i
        W(ws, f"A{r}", app, "label", bold=True); W(ws, f"B{r}", f"={R['sc_total'][app]}", "calc", "sc")
        W(ws, f"C{r}", f"=POWER(B{r},$B$34)/SUMPRODUCT(POWER($B$29:$B$32,$B$34))", "calc", "pct")
        W(ws, f"D{r}", APP_REC[app], "assump", "pct"); W(ws, f"E{r}", f"={R['sc_rank'][app]}", "calc", "int")
        W(ws, f"H{r}", APP_WHY[app], "grey", wrap=True); ws.row_dimensions[r].height = 40
    W(ws, "A33", "TOTAL (must be 100%)", "label", bold=True)
    W(ws, "C33", "=SUM(C29:C32)", "out", "pct"); W(ws, "D33", "=SUM(D29:D32)", "out", "pct")
    W(ws, "A34", "Exponent for score-implied share - ANALYST ASSUMPTION", "label"); W(ws, "B34", 3, "assump", "int")
    ws.conditional_formatting.add("D33", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
    R["app_alloc"] = {a: f"Resource_Allocation!$D${29+i}" for i, a in enumerate(APPS)}
    R["app_alloc_imp"] = {a: f"Resource_Allocation!$C${29+i}" for i, a in enumerate(APPS)}
    R["app_alloc_sum"] = "Resource_Allocation!$D$33"
    section(ws, 36, "D. Allocation of the Industrial share across submarkets", 8)
    header(ws, 37, ["Submarket", "Score (selected weights)", "Score-implied %", "RECOMMENDED %", "Rank", "", "", "Rationale"], h=30)
    for i, n in enumerate(SUBMARKETS):
        r = 38 + i
        W(ws, f"A{r}", n, "label", bold=(n.startswith("Robotics")))
        W(ws, f"B{r}", f"={R['sub_score'][n]}", "calc", "sc")
        W(ws, f"C{r}", f"=POWER(B{r},$B$46)/SUMPRODUCT(POWER($B$38:$B$44,$B$46))", "calc", "pct")
        W(ws, f"D{r}", SUB_REC[i], "assump", "pct"); W(ws, f"E{r}", f"={R['sub_rank'][n]}", "calc", "int")
        W(ws, f"H{r}", SUB_WHY[i], "grey", wrap=True); ws.row_dimensions[r].height = 30
    W(ws, "A45", "TOTAL (must be 100%)", "label", bold=True); W(ws, "C45", "=SUM(C38:C44)", "out", "pct"); W(ws, "D45", "=SUM(D38:D44)", "out", "pct")
    W(ws, "A46", "Exponent for score-implied share - ANALYST ASSUMPTION", "label"); W(ws, "B46", 6, "assump", "int")
    ws.conditional_formatting.add("D45", CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
    R["sub_alloc_sum"] = "Resource_Allocation!$D$45"
    R["rob_alloc"] = "Resource_Allocation!$D$40"
    R["rob_app_alloc"] = "Resource_Allocation!$D$40"
    W(ws, "A47", "Robotics & Physical AI share of TOTAL incremental resources (Industrial % x Robotics %)", "label", bold=True)
    W(ws, "D47", f"={R['app_alloc']['Industrial']}*D40", "out", "pct")
    R["rob_total_alloc"] = "Resource_Allocation!$D$47"
    section(ws, 49, "E. Return on incremental commercial investment (ILLUSTRATIVE - no budget is given in the research; replace with the approved figure)", 8)
    W(ws, "A51", "Incremental commercial budget, FY2027 (JPY bn) - ILLUSTRATIVE PLACEHOLDER", "label"); W(ws, "B51", 4.0, "assump", "0.0")
    header(ws, 52, ["Metric", "Base", "Bull", "Bear", "Prob-weighted", "", "", "Note"], h=24)
    names = ["Scenario probability (Industrial, from report)", "Attributable uplift: % of FY2026E Industrial revenue - ASSUMPTION", "Attributable incremental revenue (JPY bn)",
             "Gross margin (scenario)", "Incremental gross profit (JPY bn)", "Year-1 net return after commercial spend (JPY bn)", "Year-1 ROI = net return / budget",
             "Break-even incremental revenue (JPY bn) = budget / GM", "Break-even uplift as % of FY2026E Industrial revenue",
             "Cumulative gross profit over minimum product life (undiscounted, flat; JPY bn)", "Lifetime gross-profit multiple of budget (x)"]
    for k, n in enumerate(names):
        W(ws, f"A{53+k}", n, "label", wrap=True); ws.row_dimensions[53 + k].height = 28
    up = [.025, .05, .005]
    for j, key in enumerate(("base", "bull", "bear")):
        c = CL(2 + j)
        W(ws, f"{c}53", f"={R['ind_prob'][key]}", "calc", "pct")
        W(ws, f"{c}54", up[j], "assump", "pct")
        W(ws, f"{c}55", f"={c}54*{R['ind_base26']}", "calc", "bn")
        W(ws, f"{c}56", f"={R['ind_gm'][key]}", "calc", "pct")
        W(ws, f"{c}57", f"={c}55*{c}56", "calc", "bn")
        W(ws, f"{c}58", f"={c}57-$B$51", "calc", "bn")
        W(ws, f"{c}59", f"={c}58/$B$51", "calc", "pct")
        W(ws, f"{c}60", f"=$B$51/{c}56", "calc", "bn")
        W(ws, f"{c}61", f"={c}60/{R['ind_base26']}", "calc", "pct")
        W(ws, f"{c}62", f"={c}57*{R['C20']}", "calc", "bn")
        W(ws, f"{c}63", f"={c}62/$B$51", "calc", "x")
    for r, fm in ((55, "bn"), (57, "bn"), (58, "bn"), (59, "pct"), (62, "bn"), (63, "x")):
        W(ws, f"E{r}", f"=SUMPRODUCT($B$53:$D$53,B{r}:D{r})", "out", fm)
    W(ws, "E53", "=SUM(B53:D53)", "out", "pct")
    W(ws, "H54", "The report gives no revenue attributable to the programme; uplifts are assumptions. They are NOT the full 365/410/295 scenario numbers (those include macro recovery).", "note", border=False)
    W(ws, "H62", "Uses the lowest report lifecycle (7 years, C20), no decay or discounting: an upper-bound indicator, not an NPV.", "note", border=False)
    R["roi_base"], R["roi_pw"], R["breakeven"] = "Resource_Allocation!$B$59", "Resource_Allocation!$E$59", "Resource_Allocation!$B$60"
    R["roi_budget"], R["roi_uplift_base"], R["roi_rev_base"], R["roi_gm_base"] = "Resource_Allocation!$B$51", "Resource_Allocation!$B$54", "Resource_Allocation!$B$55", "Resource_Allocation!$B$56"
    R["roi_life_mult"] = "Resource_Allocation!$E$63"
    ws.conditional_formatting.add("B59:E59", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=DRED, bold=True)))
    section(ws, 65, "F. 12-month implementation roadmap (research report s.20)", 8)
    header(ws, 66, ["Phase", "Key milestones", "", "", "", "", "Owner", "Measurable outcome"], h=24)
    road = [("Days 0-90 (Nov 2026 - Jan 2027): Validate & target",
             "Mine Altium usage analytics for the top 500 robotics / physical-AI hardware developers; audit their compute and power architectures; stand up cross-functional 'Robotics Tiger Team' (analog, GaN, embedded FAEs).",
             "CCO + Altium integration office", "500 accounts scored; Tiger Team staffed; baselines set for attach rate and design-win cycle time (no baseline exists in the report)."),
            ("Months 4-6: Engage & design-in",
             "Ship RA8D1 and RZ/V2H dev kits to partners and selected startups; subsidised Reality AI access; publish pre-validated 48V cobot-joint Winning Combination (GaN + PMIC + MCU).",
             "Product + FAE leads", "Evaluations launched; reference design published; count of Altium projects selecting Renesas MCU before sales contact. Month-6 gate: review pipeline and reallocate if conversion lags."),
            ("Months 7-12: Convert & ramp",
             "FAEs convert evaluations to binding design-wins; supply chain reserves foundry capacity for prototypes and ramps; leadership reviews ASP multiplier from bundling.",
             "Strategic accounts + Supply chain", "Design-wins signed; attach rate trending to >2.5; robotics gross margin above 58% corporate baseline; channel DOI under control.")]
    for i, (p, m, o, out) in enumerate(road):
        r = 67 + i
        W(ws, f"A{r}", p, "label", wrap=True, bold=True); W(ws, f"B{r}", m, "grey", wrap=True); ws.merge_cells(f"B{r}:F{r}")
        W(ws, f"G{r}", o, "grey", wrap=True); W(ws, f"H{r}", out, "grey", wrap=True)
        ws.row_dimensions[r].height = 62
    section(ws, 71, "G. Management KPIs (research report s.21) - targets are the report's; baselines to be established", 8)
    header(ws, 72, ["KPI", "Type", "Target / threshold", "", "", "", "Unit", "Definition and use"], h=24)
    kp = [("Software-driven pipeline", "Leading", "Baseline in first 90 days", "projects", "Unique hardware projects started in Altium that natively select a Renesas MCU as primary compute BEFORE a sales rep engages."),
          ("Attachment rate", "Leading", f"={R['K20']}", "ICs per node", "Proprietary PMICs and GaN stages sold per MCU/MPU in robotics; target >2.5. Measures content per design-win."),
          ("Design-win conversion cycle", "Operational", f"={R['C14']}", "months (max)", "Days from technical evaluation to production commitment; report benchmark 12-18 months for robotics."),
          ("Channel days of inventory (DOI)", "Early warning", "Track weekly; set thresholds vs history", "days", "Real-time DOI for high-end industrial MPUs to pre-empt corrections."),
          ("Robotics gross margin vs corporate baseline", "Lagging", f"={R['K19']}", "% (min)", "Gross margin by submarket must exceed the 58% corporate baseline.")]
    for i, (k, t, tg, u, d) in enumerate(kp):
        r = 73 + i
        W(ws, f"A{r}", k, "label", bold=True); W(ws, f"B{r}", t, "grey")
        if isinstance(tg, str) and tg.startswith("="):
            fm = "pct" if "K19" in tg else "0.0"
            W(ws, f"C{r}", tg, "src", fm)
        else:
            W(ws, f"C{r}", tg, "grey", wrap=True); ws.merge_cells(f"C{r}:F{r}")
        W(ws, f"G{r}", u, "grey"); W(ws, f"H{r}", d, "grey", wrap=True); ws.row_dimensions[r].height = 30


# --------------------------------------------------------------------------------------------- SENSITIVITY
def build_sensitivity(ws):
    title(ws, "SENSITIVITY ANALYSIS - does Industrial remain #1 when assumptions change?", "All grids are live formulas (no Data Tables needed). Green = Industrial still leads; red = lead reversed. Change inputs on Strategic_Scorecard / Risk / Forecast and the grids update.", 16)
    widths(ws, [46] + [11] * 15)
    ws.freeze_panes = "B4"
    section(ws, 4, "0. Base data (links to Strategic_Scorecard Part A)", 16)
    header(ws, 5, ["Application / criterion"] + ["Growth", "Profit", "Market", "Comp. adv.", "Rev. quality", "Speed", "Strat. align", "Risk manage."] + ["WEIGHTED SCORE", "Check vs scorecard"], h=34)
    W(ws, "A6", "Weight", "label", bold=True)
    for k in range(8):
        W(ws, f"{CL(2+k)}6", f"=Strategic_Scorecard!$B${SC_FIRST+k}", "calc", "pct")
    W(ws, "J6", "=SUM(B6:I6)", "calc", "pct")
    for j, app in enumerate(APPS):
        r = 7 + j
        W(ws, f"A{r}", app, "label", bold=True)
        for k in range(8):
            W(ws, f"{CL(2+k)}{r}", f"=Strategic_Scorecard!${CL(3+j)}${SC_FIRST+k}", "calc", "sc")
        W(ws, f"J{r}", f"=SUMPRODUCT($B$6:$I$6,B{r}:I{r})", "out", "sc")
        W(ws, f"K{r}", f"=ROUND(J{r}-{R['sc_total'][app]},9)", "calc", "sc")
    W(ws, "A12", "Best competing score (excluding Industrial)", "label"); W(ws, "J12", "=MAX(J7,J9,J10)", "out", "sc")
    W(ws, "A13", "Industrial lead over best competitor (points)", "label", bold=True); W(ws, "J13", "=J8-J12", "out", "+0.00;[Red]-0.00;0.00")
    R["sens_lead"] = "Sensitivity!$J$13"
    green = PatternFill("solid", bgColor=GREEN); red = PatternFill("solid", bgColor=RED)

    def adj_block(top, kind):
        # kind 'boost' or 'zero'
        lab = "1. WEIGHT TEST A - raise one criterion's weight by +" if kind == "boost" else "2. WEIGHT TEST B - remove one criterion entirely (weight = 0, rest rescaled)"
        section(ws, top, lab + ("delta below (others rescaled proportionally)" if kind == "boost" else ""), 16)
        if kind == "boost":
            W(ws, f"A{top+1}", "Weight increase delta (percentage points) - ASSUMPTION", "label"); W(ws, f"B{top+1}", 0.10, "assump", "pct")
        header(ws, top + 2, ["Criterion changed"] + ["w:Growth", "w:Profit", "w:Market", "w:Comp.", "w:Quality", "w:Speed", "w:Align", "w:Risk"] + ["Auto", "Industrial", "Consumer", "IoT", "Winner", "Ind. lead", "Industrial #1?"], h=32)
        first = top + 3
        for k in range(8):
            r = first + k
            W(ws, f"A{r}", f"=Strategic_Scorecard!$A${SC_FIRST+k}", "label")
            wk = f"${CL(2+k)}$6"
            for j in range(8):
                src = f"{CL(2+j)}$6"
                if kind == "boost":
                    f = f"={src}+$B${top+1}" if j == k else f"={src}*(1-{wk}-$B${top+1})/(1-{wk})"
                else:
                    f = "=0" if j == k else f"={src}/(1-{wk})"
                W(ws, f"{CL(2+j)}{r}", f, "calc", "pct")
            for a in range(4):
                W(ws, f"{CL(10+a)}{r}", f"=SUMPRODUCT($B{r}:$I{r},$B${7+a}:$I${7+a})", "calc", "sc")
            W(ws, f"N{r}", f"=INDEX($J$5:$J$5,1)", "calc")
            ws[f"N{r}"].value = f'=INDEX({{"Automotive","Industrial","Consumer","IoT"}},MATCH(MAX(J{r}:M{r}),J{r}:M{r},0))'
            W(ws, f"O{r}", f"=K{r}-MAX(J{r},L{r},M{r})", "calc", "+0.00;[Red]-0.00;0.00")
            W(ws, f"P{r}", f'=IF(O{r}>0,"YES","NO")', "calc", h="center")
        ws.conditional_formatting.add(f"P{first}:P{first+7}", CellIsRule(operator="equal", formula=['"YES"'], fill=green))
        ws.conditional_formatting.add(f"P{first}:P{first+7}", CellIsRule(operator="equal", formula=['"NO"'], fill=red, font=Font(color=DRED, bold=True)))
        ws.conditional_formatting.add(f"O{first}:O{first+7}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=DRED, bold=True)))
        return first

    fa = adj_block(15, "boost")      # rows 18-25
    fb = adj_block(28, "zero")       # rows 31-38
    W(ws, "A39", "Industrial #1 in ALL 16 weight tests?", "label", bold=True)
    W(ws, "B39", f'=IF(COUNTIF(P{fa}:P{fa+7},"NO")+COUNTIF(P{fb}:P{fb+7},"NO")=0,"YES - robust to weight changes","NO - see red rows")', "out")
    W(ws, "A40", "Smallest Industrial lead across weight tests (points)", "label"); W(ws, "B40", f"=MIN(O{fa}:O{fa+7},O{fb}:O{fb+7})", "out", "+0.00;[Red]-0.00;0.00")
    R["sens_weights"], R["sens_weights_min"] = "Sensitivity!$B$39", "Sensitivity!$B$40"
    # one-way score
    section(ws, 42, "3. SCORE TEST - set Industrial's score on ONE criterion to each value 1-10 (others unchanged): lead over best competitor", 16)
    header(ws, 43, ["Industrial criterion"] + [str(s) for s in range(1, 11)] + ["Current score", "Reversal threshold", "Can reverse alone?"], h=32)
    for s in range(10):
        ws.cell(row=43, column=2 + s).value = s + 1
    for k in range(8):
        r = 44 + k
        W(ws, f"A{r}", f"=Strategic_Scorecard!$A${SC_FIRST+k}", "label")
        for s in range(10):
            c = CL(2 + s)
            W(ws, f"{c}{r}", f"=$J$8+${CL(2+k)}$6*({c}$43-${CL(2+k)}$8)-$J$12", "calc", "+0.00;[Red]-0.00;0.00")
        W(ws, f"L{r}", f"={CL(2+k)}$8", "calc", "sc")
        W(ws, f"M{r}", f"=L{r}-$J$13/{CL(2+k)}$6", "out", "0.0")
        W(ws, f"N{r}", f'=IF(M{r}<1,"No (needs score <1)","Yes if score <= "&TEXT(M{r},"0.0"))', "calc")
    ws.conditional_formatting.add("B44:K51", CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add("B44:K51", CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=red))
    # two-way comp adv x align
    section(ws, 53, "4. SUBJECTIVE-SCORE TEST - cut Industrial 'Competitive advantage' (down) and 'Strategic alignment' (across) simultaneously: lead over best competitor", 16)
    header(ws, 54, ["Cut in Competitive advantage \\ cut in Strategic alignment"] + [str(i) for i in range(0, 7)], h=32)
    for i in range(7):
        ws.cell(row=54, column=2 + i).value = i
    for a in range(7):
        r = 55 + a
        W(ws, f"A{r}", a, "assump", "0", h="center")
        for c_ in range(7):
            c = CL(2 + c_)
            W(ws, f"{c}{r}", f"=$J$8-$E$6*$A{r}-$H$6*{c}$54-$J$12", "calc", "+0.00;[Red]-0.00;0.00")
    ws.conditional_formatting.add("B55:H61", CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add("B55:H61", CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=red))
    # risk weight x shock
    section(ws, 63, "5. RISK TEST - raise the weight on risk manageability (down) AND worsen Industrial's risk score by 'shock' points (across): lead over best competitor", 16)
    header(ws, 64, ["Risk weight \\ Industrial risk-score shock"] + [str(i) for i in range(0, 5)], h=32)
    for i in range(5):
        ws.cell(row=64, column=2 + i).value = i
    for a, x in enumerate((.05, .10, .15, .20, .30, .40)):
        r = 65 + a
        W(ws, f"A{r}", x, "assump", "pct", h="center")
        for c_ in range(5):
            c = CL(2 + c_)
            ind = f"(($J$8-$I$6*$I$8)*(1-$A{r})/(1-$I$6)+$A{r}*MAX(1,$I$8-{c}$64))"
            oth = lambda rr: f"(($J${rr}-$I$6*$I${rr})*(1-$A{r})/(1-$I$6)+$A{r}*$I${rr})"
            W(ws, f"{c}{r}", f"={ind}-MAX({oth(7)},{oth(9)},{oth(10)})", "calc", "+0.00;[Red]-0.00;0.00")
    ws.conditional_formatting.add("B65:F70", CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add("B65:F70", CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=red))
    # revenue growth haircut on Industrial growth score
    section(ws, 72, "6. FINANCIAL TEST - Industrial FY2027 gross profit (JPY bn) = FY2026E base x (1+growth) x gross margin; green if >= probability-weighted expected GP", 16)
    W(ws, "A73", "FY2026E Industrial base (JPY bn)", "label"); W(ws, "B73", f"={R['ind_base26']}", "calc", "bn")
    W(ws, "C73", "Expected GP:", "label", border=False); W(ws, "D73", f"={R['ind_exp_gp']}", "out", "bn")
    header(ws, 74, ["Gross margin \\ revenue growth vs FY2026E"] + [str(i) for i in range(7)], h=32)
    gs = [-.10, -.05, 0.0, .05, .10, .15, .20]
    for j, g in enumerate(gs):
        W(ws, f"{CL(2+j)}74", g, "assump", "pct", h="center")
        ws[f"{CL(2+j)}74"].font = Font(name=FONT, size=10, bold=True, color="0000FF")
    for a, gm in enumerate((.53, .55, .57, .59, .61, .63, .65)):
        r = 75 + a
        W(ws, f"A{r}", gm, "assump", "pct", h="center")
        for j in range(7):
            c = CL(2 + j)
            W(ws, f"{c}{r}", f"=$B$73*(1+{c}$74)*$A{r}", "calc", "bn")
    ws.conditional_formatting.add("B75:H81", FormulaRule(formula=["B75>=$D$73"], fill=green))
    ws.conditional_formatting.add("B75:H81", FormulaRule(formula=["B75<$D$73"], fill=red))
    # probabilities
    section(ws, 83, "7. DEMAND / PROBABILITY TEST - Industrial expected FY2027 revenue (JPY bn) as Bear (down) and Bull (across) probabilities change; Base = remainder", 16)
    header(ws, 84, ["Bear probability \\ Bull probability"] + [str(i) for i in range(5)], h=32)
    for j, p in enumerate((.10, .15, .20, .25, .30)):
        W(ws, f"{CL(2+j)}84", p, "assump", "pct", h="center")
        ws[f"{CL(2+j)}84"].font = Font(name=FONT, size=10, bold=True, color="0000FF")
    for a, pb in enumerate((.10, .20, .30, .40, .50)):
        r = 85 + a
        W(ws, f"A{r}", pb, "assump", "pct", h="center")
        for j in range(5):
            c = CL(2 + j)
            W(ws, f"{c}{r}", f"=IF(1-$A{r}-{c}$84<0,\"n/a\",(1-$A{r}-{c}$84)*{R['ind_rev']['base']}+{c}$84*{R['ind_rev']['bull']}+$A{r}*{R['ind_rev']['bear']})", "calc", "bn")
    ws.conditional_formatting.add("B85:F89", FormulaRule(formula=[f"AND(ISNUMBER(B85),B85>={R['ind_base26']})"], fill=green))
    ws.conditional_formatting.add("B85:F89", FormulaRule(formula=[f"AND(ISNUMBER(B85),B85<{R['ind_base26']})"], fill=red))
    W(ws, "A90", "Green = expected FY2027 Industrial revenue above the FY2026E base (growth); red = decline.", "note", border=False)
    # conversion timing
    section(ws, 92, "8. CONVERSION-TIMELINE TEST - FY2027 revenue attributable to the programme (JPY bn): delay in design-win conversion (down) x share of pipeline converting (across)", 16)
    W(ws, "A93", "Base-case attributable revenue if on time and 100% converts (JPY bn)", "label"); W(ws, "B93", f"={R['roi_rev_base']}", "calc", "bn")
    header(ws, 94, ["Delay (months) \\ pipeline conversion"] + [str(i) for i in range(6)], h=32)
    for j, p in enumerate((.5, .6, .7, .8, .9, 1.0)):
        W(ws, f"{CL(2+j)}94", p, "assump", "pct", h="center")
        ws[f"{CL(2+j)}94"].font = Font(name=FONT, size=10, bold=True, color="0000FF")
    for a, dl in enumerate((0, 3, 6, 9, 12)):
        r = 95 + a
        W(ws, f"A{r}", dl, "assump", "0", h="center")
        for j in range(6):
            c = CL(2 + j)
            W(ws, f"{c}{r}", f"=$B$93*{c}$94*MAX(0,(12-$A{r})/12)", "calc", "bn")
    header(ws, 101, ["Year-1 ROI (net of budget) \\ pipeline conversion"] + [str(i) for i in range(6)], h=32)
    for j in range(6):
        W(ws, f"{CL(2+j)}101", f"={CL(2+j)}94", "calc", "pct", h="center")
        ws[f"{CL(2+j)}101"].font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws[f"{CL(2+j)}101"].fill = PatternFill("solid", fgColor=NAVY)
    for a in range(5):
        r = 102 + a
        W(ws, f"A{r}", f"=A{95+a}", "calc", "0", h="center")
        for j in range(6):
            c = CL(2 + j)
            W(ws, f"{c}{r}", f"=({c}{95+a}*{R['roi_gm_base']}-{R['roi_budget']})/{R['roi_budget']}", "calc", "pct")
    ws.conditional_formatting.add("B102:G106", CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add("B102:G106", CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=red))
    W(ws, "A107", "Key reading: a 12-18 month robotics design-win cycle means FY2027 revenue from NEW wins is limited; Year-1 ROI turns negative with a 6+ month delay or <70% conversion. The case rests on lifetime value, not Year-1 payback.", "note", border=False)
    # submarket robustness
    section(ws, 109, "9. SUBMARKET TEST (Part B) - cut Robotics 'Competitive advantage' (down) and 'Strategic alignment' (across): Robotics lead over best other submarket (selected weights)", 16)
    header(ws, 110, ["Cut in Competitive advantage \\ cut in Strategic alignment"] + [str(i) for i in range(6)], h=32)
    for i in range(6):
        ws.cell(row=110, column=2 + i).value = i
    for a in range(6):
        r = 111 + a
        W(ws, f"A{r}", a, "assump", "0", h="center")
        for c_ in range(6):
            c = CL(2 + c_)
            W(ws, f"{c}{r}", f"=Strategic_Scorecard!$H$48-Strategic_Scorecard!$E$40*$A{r}-Strategic_Scorecard!$E$44*{c}$110-MAX(Strategic_Scorecard!$F$48:$G$48,Strategic_Scorecard!$I$48:$L$48)", "calc", "+0.00;[Red]-0.00;0.00")
    ws.conditional_formatting.add("B111:G116", CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add("B111:G116", CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=red))
    W(ws, "A117", "Robotics remains first only while its two subjective scores are not cut by more than ~3-4 points combined: AI data centres is the nearest rival and would overtake.", "note", border=False)
    R["sub_haircut_cell"] = "Sensitivity!$B$111"
    section(ws, 119, "10. SUMMARY - what could reverse the recommendation?", 16)
    rows = [("Application level: Industrial lead (points)", f"={R['sens_lead']}".replace("$J$13", "$J$13")),
            ("Application level: weight tests", "=B39"),
            ("Application level: smallest Industrial lead across weight tests", "=B40"),
            ("Submarket level: Robotics lead - selected weights (points)", f"={R['sub_lead']}"),
            ("Submarket level: winner under report / assignment / custom weights", f'={"Strategic_Scorecard!$B$57"}&" | "&{"Strategic_Scorecard!$B$58"}&" | "&{"Strategic_Scorecard!$B$59"}')]
    for i, (l, f) in enumerate(rows):
        W(ws, f"A{120+i}", l, "label"); W(ws, f"B{120+i}", f, "out", "+0.00;[Red]-0.00;0.00")
    text_rows(ws, 126, ["Reversal conditions (formula-driven thresholds in section 3 and grids 4-9): (1) Industrial 'Revenue growth' or 'Competitive advantage' judged materially weaker; (2) risk weight raised to >=30% with a >=2-point Industrial risk shock;",
                        "(3) Industrial scenario probabilities shift toward Bear (e.g. bear >=40%); (4) robotics design-win conversion slips beyond ~6 months with <70% conversion (Year-1 ROI negative, lifetime case intact); (5) Altium fails to generate silicon pull-through (Robotics alignment/competitive scores cut by 4+)."])
    R["sens_verdict_cell"] = "Sensitivity!$B$39"


# --------------------------------------------------------------------------------------------- VALIDATION
FLAGS = [
    ("F-01", "No assignment dataset supplied; only a research report.", "All files", "No actual application-level dataset exists; any 'raw data' is research-derived.", "Built a clearly labelled research-derived dataset (Raw_Data). Application-level Consumer/IoT splits and margins are assumptions/proxies; workbook is built so a real dataset can replace Raw_Data.", "High"),
    ("F-02", "FY2026 shown as 'annualised est.' next to FY2023-25 actuals; report calls FY2026 'latest available' while research cut-off is 2 Nov 2026 (year not finished).", "s.7 table", "Mixing actuals and estimates distorts growth: +14% FY2026E growth is partly estimation method.", "FY2026E labelled '(est.)' everywhere; CAGR shown both on actuals only (FY23-25) and including FY26E; no post-cutoff verification was performed.", "Medium"),
    ("F-03", "Pure Industrial low-end 282 bn is inconsistent with 40% x 671.8 = 268.7 bn.", "s.4 text vs table", "Understates uncertainty band (true band ~269-302 bn, +/-6%).", "Recomputed in Revenue_Analysis s.6; point estimate 285.5 reconciles.", "Medium"),
    ("F-04", "I&I operating margin (27.2%) is BELOW consolidated (29.3%) while gross margin is above (60.6% vs 57.6%): implied Automotive operating margin is ~31.5%.", "s.7, s.9", "Contradicts 'Industrial offers superior structural profitability' at operating level; advantage is gross-margin only (heavier R&D/SG&A).", "Derived Automotive margins shown in Profitability_Quality s.2; scorecard profitability scores kept close (Industrial 7 vs Automotive 6).", "High"),
    ("F-05", "Report's weighted scorecard totals do not reconcile to its own scores and weights (e.g. Robotics 8.40 reported vs 8.20 recomputed; Factory Automation 6.60 vs 6.85).", "s.14", "Headline margin of Robotics over AI data centres is 0.85 not 1.05 points.", "Recomputed with SUMPRODUCT in Strategic_Scorecard Part B; ranking unchanged.", "Medium"),
    ("F-06", "Report scorecard ranks 7 Industrial submarkets on 9 criteria; the assignment requires the four applications on 8 criteria.", "s.14", "Report cannot answer the Automotive vs Industrial vs Consumer vs IoT question directly.", "Built Part A (four applications, assignment weights) and kept Part B (submarkets) with weight-set selector.", "Info"),
    ("F-07", "Robotics market growth quoted as '>30%' (s.6) and '>25%' (s.11); market sized only by a 2035 end value (USD 58bn) with no base year; figure is a total Physical-AI market from a third-party site, not Renesas-addressable semiconductor content.", "s.6, s.11", "CAGR cannot be derived; growth may be overstated; market CAGR is not company CAGR.", "Treated as qualitative evidence; CAGR not computed (Growth_Opportunity s.1).", "High"),
    ("F-08", "No Renesas robotics / physical-AI revenue is disclosed. The 365/410/295 bn FY2027 scenarios are for ALL pure Industrial, but narrative attributes outcomes to robotics. 12-18 month design-win cycles imply little FY2027 revenue from NEW wins.", "s.17, s.15", "Cannot attribute scenario revenue to the incremental robotics programme.", "Attributable uplift modelled separately with explicit assumptions (Resource_Allocation s.E) and stress-tested (Sensitivity s.8).", "High"),
    ("F-09", "Altium (USD 263m revenue, ~3% of group) is credited with making revenue 'largely immune to cyclicality', an 'unassailable workflow monopoly' and an 'insurmountable moat'. Altium is a vendor-neutral tool; no evidence of silicon pull-through is provided.", "s.8, s.10, s.15", "Overstates revenue-quality benefit; thesis is unproven.", "Revenue-quality scores do not credit Altium beyond its measured size; Robotics scores cut in sensitivity; attach-rate KPI used to test.", "High"),
    ("F-10", "Board question 12 treats 'USD 3bn' Timing proceeds as deployable; sale is 50% cash / 50% SiTime stock (about USD 1.5bn cash).", "s.7, s.22", "Funding capacity for the robotics push may be overstated.", "Noted; no Timing proceeds used in the model; budget is an explicit placeholder.", "Medium"),
    ("F-11", "'China ~30% of revenue' rests on an aggregator source [15]; 'no single customer >10%' is stated without verification; both sit uneasily with narrative of diversification.", "s.8, s.19, s.22", "Geopolitical and concentration risk may be mis-scored.", "Flagged as Report claim; used only qualitatively in risk scoring.", "Medium"),
    ("F-12", "Report claims I&I 'structural resilience', but data show I&I fell 16.7% in FY2024 while Automotive rose 1.0%; I&I YoY volatility exceeds Automotive's.", "s.7", "Revenue-stability advantage of Industrial/I&I is not demonstrated.", "Measured volatility proxy added (Profitability_Quality); Industrial stability scored 5 vs Automotive 6.", "Medium"),
    ("F-13", "Source quality: references include a job description document [10], an aggregator [15], a stock-moat blog [6], consulting-site summary [12] and market-research vendors for TAM.", "Works cited", "Evidence authority is uneven; Renesas filings are the only high-authority sources.", "Source authority recorded in Sources_Assumptions; third-party market data labelled as such.", "Medium"),
    ("F-14", "Language such as 'definitive, unequivocal', 'unassailable', 'insurmountable' is not supported by scorecard (which mixes judgement scores) or by data.", "s.1, s.10, s.15", "Over-confident framing.", "Independent scoring and sensitivity; recommendation stated with its reversal conditions.", "Medium"),
]


def build_validation(ws):
    title(ws, "DATA CLEANING AND VALIDATION", "Documents dataset availability, quality checks (live formulas), report inconsistencies, transformations and a reliability summary. Nothing was silently removed.", 7)
    widths(ws, [8, 62, 18, 28, 12, 78, 10])
    ws.freeze_panes = "A4"
    section(ws, 4, "A. Dataset availability assessment", 7)
    avail = ["Files provided: (1) Renesas Industrial Strategy Analysis.pdf, (2) the same report as Markdown. No Excel/CSV assignment dataset was provided.",
             "Conclusion: an ACTUAL assignment dataset is NOT available. The workbook uses a research-derived dataset (Raw_Data) built only from figures in the report; nothing is fabricated.",
             "Coverage: Automotive and I&I segment revenue FY2023-FY2025 (+FY2026 annualised est.), consolidated margins, I&I FY2025 margins, market estimates, scenario inputs. No customer, product, geography or application-level margin data.",
             "Gaps: Industrial/Consumer/IoT revenue and profit are not reported; Consumer and IoT cannot be separated. Treatment: estimate (Industrial), illustrative split (Consumer/IoT), proxies (margins) - all labelled.",
             "Sufficiency: sufficient for a transparent directional recommendation; NOT sufficient for audited application-level P&L. If the original dataset arrives: paste it into Raw_Data, repoint Revenue_Analysis row links, re-run checks here."]
    text_rows(ws, 5, avail, 1)
    section(ws, 11, "B. Quality checks (live formulas)", 7)
    header(ws, 12, ["#", "Check", "Result", "Expectation", "Status", "Comment", ""], h=24)
    ids = R["raw_ids_range"]
    checks = [
        ("Segment sum vs total, FY2023 (residual as % of revenue)", "=Raw_Data!C17", "|x| < 1%", '=IF(ABS(C{r})<0.01,"PASS","FLAG")', "pct", "Auto + I&I leave a small 'Other' residual; shown explicitly."),
        ("Segment sum vs total, FY2024", "=Raw_Data!D17", "|x| < 1%", '=IF(ABS(C{r})<0.01,"PASS","FLAG")', "pct", ""),
        ("Segment sum vs total, FY2025", "=Raw_Data!E17", "|x| < 1%", '=IF(ABS(C{r})<0.01,"PASS","FLAG")', "pct", "9.7 bn residual (0.7%)."),
        ("Segment sum vs total, FY2026E", "=Raw_Data!F17", "|x| < 1%", '=IF(ABS(C{r})<0.01,"PASS","FLAG")', "pct", ""),
        ("Application revenue components reconcile to reported total", "=SUMPRODUCT(ABS(Revenue_Analysis!B13:E13))", "= 0", '=IF(C{r}=0,"PASS","FLAG")', "0.000", "Industrial+Consumer+IoT = I&I exactly: no double counting."),
        ("Missing values in Table 1 (blank cells)", "=COUNTBLANK(Raw_Data!C7:F11)", "= 0", '=IF(C{r}=0,"PASS","FLAG")', "0", ""),
        ("Duplicate IDs in key-value table", f"=SUMPRODUCT((COUNTIF({ids},{ids})>1)*1)", "= 0", '=IF(C{r}=0,"PASS","FLAG")', "0", ""),
        ("Unit consistency (JPY bn vs USD bn vs USD m)", "=\"Labelled\"", "Units labelled, never mixed", '="PASS"', None, "USD figures (market sizes, Altium) are never added to JPY revenue. Implied FX shown only for context."),
        ("Pure Industrial FY2025 low-end: calc less report (JPY bn)", "=Revenue_Analysis!F57", "|x| <= 1", '=IF(ABS(C{r})<=1,"PASS","FLAG")', "bn", "Report's 282 vs 268.7 calculated (F-03)."),
        ("Pure Industrial FY2025 point: calc less report", "=Revenue_Analysis!F58", "|x| <= 1", '=IF(ABS(C{r})<=1,"PASS","FLAG")', "bn", ""),
        ("Pure Industrial FY2025 high-end: calc less report", "=Revenue_Analysis!F59", "|x| <= 1", '=IF(ABS(C{r})<=1,"PASS","FLAG")', "bn", ""),
        ("Implied FY2026E Industrial share of I&I within 40-45%", "=Revenue_Analysis!E15", "40% to 45%", f'=IF(AND(C{{r}}>={R["K03"]},C{{r}}<={R["K04"]}),"PASS","FLAG")', "pct", "Report's 320 / annualised I&I 772."),
        ("Outlier scan: |YoY| > 10% observations (Automotive & I&I)", "=SUMPRODUCT((ABS(Revenue_Analysis!C32:E32)>0.1)*1)+SUMPRODUCT((ABS(Revenue_Analysis!C37:E37)>0.1)*1)", "= 0", '=IF(C{r}=0,"PASS","FLAG")', "0", "Genuine cyclical swings (FY2024 correction; FY2025 auto drop; FY2026E rebound) - RETAINED, not removed."),
        ("Report scorecard weights sum to 100%", "=Raw_Data!C" + str(R["sub_total_row"]), "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Report scorecard totals reconcile to its own scores (count of mismatches)", f"={R['sub_flags']}", "0", '=IF(C{r}=0,"PASS","FLAG")', "0", "See F-05."),
        ("Industrial scenario probabilities (report) sum to 100%", f"={R['K16']}+{R['K17']}+{R['K18']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("All forecast scenario probabilities sum to 100%", f"={R['fc_prob_check']}", "PASS", '=IF(C{r}="PASS","PASS","FLAG")', None, ""),
        ("Strategic scorecard (Part A) weights sum to 100%", f"={R['sc_wsum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Risk factor weights sum to 100%", f"={R['risk_wsum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Revenue-quality weights sum to 100%", f"={R['q_wsum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Market-opportunity weights sum to 100%", f"={R['opp_wsum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Submarket weight sets (report/assignment/custom/selected) all 100%", "=SUMPRODUCT(" + "*".join(f"(ROUND({a},6)=1)" for a in R["sub_wsums"]) + ")", "1 (all four true)", '=IF(C{r}=1,"PASS","FLAG")', "0", ""),
        ("Resource allocation - report (must be 100%)", f"={R['alloc_rep_sum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Resource allocation - recommended (must be 100%)", f"={R['alloc_rec_sum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Application allocation (must be 100%)", f"={R['app_alloc_sum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Industrial submarket allocation (must be 100%)", f"={R['sub_alloc_sum']}", "100%", '=IF(ROUND(C{r},6)=1,"PASS","FLAG")', "pct", ""),
        ("Margin proxies reconcile to I&I gross profit (diff, JPY bn)", "=Profitability_Quality!F36", "0", '=IF(ABS(C{r})<0.001,"PASS","FLAG")', "0.000", ""),
        ("Margin proxies reconcile to I&I operating profit (diff, JPY bn)", "=Profitability_Quality!F37", "0", '=IF(ABS(C{r})<0.001,"PASS","FLAG")', "0.000", ""),
        ("Scorecard contributions equal SUMPRODUCT (sum abs diff)", "=SUMPRODUCT(ABS(Strategic_Scorecard!C32:F32))", "0", '=IF(C{r}<0.000001,"PASS","FLAG")', "0.000", ""),
        ("Dashboard total revenue equals Revenue_Analysis total (diff)", "=ROUND(Dashboard!B14-Revenue_Analysis!D11,6)", "0", '=IF(C{r}=0,"PASS","FLAG")', "0.000", "Dashboard is formula-linked."),
        ("Altium revenue as % of group revenue (context)", f"={R['C01']}*{R['FX']}/1000/Revenue_Analysis!D12", "info", '="INFO"', "pct", "~3%: too small to transform group revenue quality (F-09). USD 263m x implied FX."),
        ("Probability-weighted Industrial revenue vs base case (JPY bn)", f"={R['ind_exp_rev']}-{R['ind_rev']['base']}", "info", '="INFO"', "bn", "Expected 360 vs base 365: scenario skew is mildly negative."),
    ]
    r = 13
    for i, (n, f, ex, st, fm, cm) in enumerate(checks):
        W(ws, f"A{r}", i + 1, "grey", "0", h="center"); W(ws, f"B{r}", n, "label", wrap=True)
        W(ws, f"C{r}", f, "calc", fm, h="center"); W(ws, f"D{r}", ex, "grey", h="center")
        W(ws, f"E{r}", st.replace("{r}", str(r)).replace("{{r}}", str(r)), "calc", h="center", bold=True)
        W(ws, f"F{r}", cm, "grey", wrap=True)
        ws.row_dimensions[r].height = 28
        r += 1
    last = r - 1
    ws.conditional_formatting.add(f"E13:E{last}", CellIsRule(operator="equal", formula=['"PASS"'], fill=PatternFill("solid", fgColor=GREEN), font=Font(color=DGREEN, bold=True)))
    ws.conditional_formatting.add(f"E13:E{last}", CellIsRule(operator="equal", formula=['"FLAG"'], fill=PatternFill("solid", fgColor=RED), font=Font(color=DRED, bold=True)))
    W(ws, f"B{r+1}", "Checks PASS / FLAG / INFO", "label", bold=True)
    W(ws, f"C{r+1}", f'=COUNTIF(E13:E{last},"PASS")&" / "&COUNTIF(E13:E{last},"FLAG")&" / "&COUNTIF(E13:E{last},"INFO")', "out", h="center")
    W(ws, f"B{r+2}", "Number of FLAGs that are explained report/data issues (expected)", "label"); W(ws, f"C{r+2}", f'=COUNTIF(E13:E{last},"FLAG")', "out", "0", h="center")
    R["val_summary"] = f"Data_Validation!$C${r+1}"
    R["val_flags"] = f"Data_Validation!$C${r+2}"
    R["val_last"] = last
    section(ws, r + 4, "C. Dataset reliability summary", 7)
    rel = [("Source data (reported*)", "MODERATE", "Report cites Renesas filings but figures were not re-verified; FY2026 is an annualised estimate."),
           ("Industrial revenue", "LOW-MODERATE", "Estimate: 40-45% share assumption of I&I; +/-6% band around 285.5; no standalone disclosure."),
           ("Consumer / IoT split", "LOW", "Illustrative assumption; no data."),
           ("Application margins", "LOW-MODERATE", "Automotive derived; others proxies reconciled to I&I."),
           ("Market growth data", "LOW-MODERATE", "Third-party TAM figures (total markets), inconsistent CAGRs, no robotics base year."),
           ("Scorecards / risk scores", "SUBJECTIVE", "Analyst judgement anchored to report evidence; stress-tested in Sensitivity."),
           ("OVERALL: sufficient for a directional, transparent recommendation; not for audited application P&L", "MODERATE-LOW", "Replace Raw_Data with the real assignment dataset when available.")]
    header(ws, r + 5, ["", "Dimension", "Reliability", "", "", "Reasoning", ""], h=22)
    for i, (d, lv, rs) in enumerate(rel):
        rr = r + 6 + i
        W(ws, f"B{rr}", d, "label", wrap=True, bold=(i == 6)); W(ws, f"C{rr}", lv, "assump", h="center"); W(ws, f"F{rr}", rs, "grey", wrap=True)
        ws.row_dimensions[rr].height = 30
    s0 = r + 6 + len(rel) + 1
    section(ws, s0, "D. Report inconsistency and claim log (flagged, not silently accepted)", 7)
    header(ws, s0 + 1, ["ID", "Issue", "Where in report", "Why it matters", "Severity", "Treatment in this workbook", ""], h=24)
    for i, (id_, iss, wh, why, tr, sev) in enumerate(FLAGS):
        rr = s0 + 2 + i
        W(ws, f"A{rr}", id_, "grey"); W(ws, f"B{rr}", iss, "label", wrap=True); W(ws, f"C{rr}", wh, "grey", wrap=True)
        W(ws, f"D{rr}", why, "grey", wrap=True); W(ws, f"E{rr}", sev, "flag" if sev == "High" else "grey", h="center"); W(ws, f"F{rr}", tr, "grey", wrap=True)
        ws.row_dimensions[rr].height = 76
    t0 = s0 + 2 + len(FLAGS) + 1
    section(ws, t0, "E. Data transformations and cleaning decisions", 7)
    header(ws, t0 + 1, ["#", "Transformation / decision", "Applies to", "", "", "Rationale", ""], h=22)
    tr = [("Retained every observation; no rows deleted or overwritten", "All tables", "Source values remain in Raw_Data exactly as presented."),
          ("Added calculated 'Other / unallocated' residual", "Table 1", "Components do not sum exactly to total; residual disclosed instead of forced."),
          ("Estimated Industrial = I&I x 42.5% (FY23-25 constant share); FY2026E = report's 320", "Revenue_Analysis", "Only the FY2025 share is given; constant-share assumption flagged (growth equals I&I by construction)."),
          ("Illustrative Consumer / IoT split of the non-industrial remainder (40/60)", "Revenue_Analysis", "No source; editable yellow input; reconciles to I&I."),
          ("Margin proxies adjusted so Industrial + Consumer + IoT equal reported I&I margin", "Profitability_Quality", "Prevents double counting / unreconciled margins."),
          ("Converted risk scale to manageability = 11 - risk", "Scorecard", "Ensures higher score = more attractive for every criterion."),
          ("Percentages stored as decimals with % formatting; money in JPY bn with one decimal", "All sheets", "Consistent precision; avoids unit confusion."),
          ("Recomputed report scorecard totals via SUMPRODUCT", "Strategic_Scorecard Part B", "Report totals do not reconcile (F-05).")]
    for i, (a, b, c) in enumerate(tr):
        rr = t0 + 2 + i
        W(ws, f"A{rr}", i + 1, "grey", "0", h="center"); W(ws, f"B{rr}", a, "label", wrap=True); W(ws, f"C{rr}", b, "grey", wrap=True)
        W(ws, f"F{rr}", c, "grey", wrap=True); ws.row_dimensions[rr].height = 30


# --------------------------------------------------------------------------------------------- SOURCES
REFS = [
    "Regarding Acquisition of Stock of Altium Limited - Renesas|https://www.renesas.com/en/about/newsroom/regarding-acquisition-stock-altium-limited",
    "Renesas to Acquire PCB Design Software Leader Altium|https://www.renesas.com/en/about/newsroom/renesas-acquire-pcb-design-software-leader-altium-make-electronics-design-accessible-broader-market",
    "Renesas Announces Expected Loss Resulting from Signing Restructuring Support Agreement (Wolfspeed)|https://www.renesas.com/en/about/newsroom/renesas-announces-expected-loss-resulting-signing-restructuring-support-agreement-wolfspeed",
    "Analyst Briefing on Wolfspeed Restructuring Support Agreement|https://www.renesas.com/document/ppt/wolfspeed-restructuring-support-agreement-presentation-and-qa",
    "Renesas Reports Financial Results for the Year Ended December 31, 2025|https://www.renesas.com/document/rep/earnings-report-year-ended-december-31-2025",
    "NXP Semiconductors N.V. Business & Moat Analysis (2026) - koalagains|https://koalagains.com/stocks/XETRA/VNX/business-and-moat",
    "RA8P1 - 1GHz Arm Cortex-M85 and Ethos-U55 NPU based AI MCU - Renesas|https://www.renesas.com/en/products/ra8p1",
    "Can Your Vision AI Solution Keep Up with Cortex-M85? - Renesas|https://www.renesas.com/en/blogs/can-your-vision-ai-solution-keep-cortex-m85",
    "DA9062 - PMIC Designed for Applications Requiring up to 8.5A|https://www.renesas.com/en/products/da9062",
    "JD_Summer Placements_Renesas Electronics.docx (local document; no URL - job description, low authority)|n/a",
    "POWER 2024 Capital Market Day - Renesas|https://www.renesas.com/en/document/ppt/power-2024-capital-market-day",
    "Renesas Strategy and Business Model - Umbrex|https://umbrex.com/resources/company-profiles/renesas-electronics-corp/",
    "Renesas Announces More than 35 Winning Combinations|https://www.renesas.com/en/about/newsroom/renesas-announces-more-35-winning-combinations-featuring-both-dialog-and-renesas-products",
    "Renesas 365 Launches: From Inspired Vision to Reality|https://www.renesas.com/en/blogs/renesas-365-launches-inspired-vision-reality",
    "Renesas Electronics (6723) Revenue Breakdown Overview - Futu (aggregator)|https://www.futunn.com/en/stock/6723-JP/financials-revenue",
    "Press Release - Renesas - DiscoverEE|https://www.discoveree.io/morepressrelease/renesas",
    "EEAWARDS2024|https://site.eettaiwan.com/events/eeaward2024/index_en.html",
    "Multi-Channel Power Management ICs (PMICs) - Renesas|https://www.renesas.com/en/products/power-management/multi-channel-power-management-ics-pmics",
    "Renesas Strengthens Power Leadership with New GaN FETs for AI Data Centers|https://www.renesas.com/en/about/newsroom/renesas-strengthens-power-leadership-new-gan-fets-high-density-power-conversion-ai-data-centers",
    "Renesas Unveils 650 V GaN FETs for Power Conversion - EEPower|https://eepower.com/news/renesas-unveils-650v-gan-fets-for-power-conversion/",
    "What Is 800V HVDC: AI Data Centers Power Architecture - COBTEL (URL wrapped in report; verify)|https://www.cobtel.com/info/what-is-nvidia-800v-hvdc-the-power-architectu-103594125.html",
    "CL8040 - Wi-Fi 6 Concurrent Dual Band 4T4R PCIe Chip - Renesas|https://www.renesas.com/en/products/cl8040",
    "Industrial Control & Factory Automation Market worth $435.24 billion - MarketsandMarkets|https://www.marketsandmarkets.com/core/PressReleases/factory-automation.asp",
    "Factory Automation Market Size & Growth Trends 2031 - Mordor Intelligence|https://www.mordorintelligence.com/industry-reports/global-factory-automation-and-industrial-controls-market-industry",
    "Power Semiconductor Market Size, Share & Forecast 2034 - IMARC|https://www.imarcgroup.com/power-semiconductor-market",
    "Physical AI Market Size, Share | Growth [2035] - Astute Analytica|https://www.astuteanalytica.com/industry-report/physical-ai-market",
    "AI Infrastructure Market Size, Growth and Forecast Report 2030 - TechSci|https://www.techsciresearch.com/report/ai-infrastructure-market/30524.html",
    "Renesas unveils next-gen GaN FETs for data centers - EDN|https://www.edn.com/renesas-unveils-next-gen-gan-fets-for-data-centers/",
    "Analog Integrated Circuit (IC) Industry - Mordor Intelligence|https://www.mordorintelligence.com/industry-reports/analog-integrated-circuit-market",
    "Analog Integrated Circuit Market Size, Growth, Forecast Report 2035 - MRFR|https://www.marketresearchfuture.com/reports/analog-integrated-circuit-market-30156",
    "Financial Report 2023 - Renesas|https://www.renesas.com/en/document/rep/financial-report-2023",
    "2026 2Q Presentation Material - Renesas|https://www.renesas.com/en/document/ppt/2026-2q-presentation-material",
    "Renesas Q2 2026 slides: revenue surges 25%, operating margin - Investing.com|https://www.investing.com/news/company-news/renesas-q2-2026-slides-revenuesurges-25-operating-margin-hits-33-93CH-4827015",
    "Wolfspeed Nears Bankruptcy; Renesas to Record 250B Yen Loss - smbom|https://www.smbom.com/news/40327",
    "Renesas Electronics (RNECF) Earnings Date - MarketBeat|https://www.marketbeat.com/earnings/reports/2026-10-29-renesas-electronics-costock/",
    "SiTime to Acquire Renesas' Timing Business|https://www.sitime.com/company/newsroom/press-release/sitime-acquire-renesastiming-business",
    "Financial Report 2024 - Renesas|https://www.renesas.com/en/document/rep/financial-report-2024",
    "Financial Report 2022 - Renesas|https://www.renesas.com/en/document/rep/financial-report-2022",
    "Renesas Financial Report 2025|https://www.renesas.com/en/document/rep/financial-report-2025",
    "Texas Instruments (TXN) financial and ratio analysis: Q2 2026 filings - sec-api|https://sec-api.io/insights/texas-instruments-txn-financial-and-ratio-analysis-q2-2026-filings-analog-sector-comparison-and-scenarios-to-2029",
    "Analog Devices Reports Strong Fourth Quarter and Fiscal 2025|https://investor.analog.com/news-releases/news-release-details/analog-devices-reports-strong-fourth-quarter-and-fiscal-2025",
    "Company presentation - Infineon Technologies FY2026 Q2|https://www.infineon.com/assets/row/public/documents/corporate/fy2026-q2-web-en-new.pdf",
    "Robotics: Technologies and Global Markets - BCC Research|https://www.bccresearch.com/market-research/engineering/robotics.html",
    "Renesas Delivers New RA8 MCU Group Targeting Graphic Display|https://www.renesas.com/en/about/newsroom/renesas-delivers-new-ra8-mcu-group-targeting-graphic-display-solutions-and-voicevision-multimodal-ai",
    "4Q/Full-Year 2025 Presentation - Renesas|https://www.renesas.com/en/document/ppt/2025-full-year-presentation-material",
    "Renesas Partners with Indian Government to Drive Innovation|https://www.renesas.com/en/about/newsroom/renesas-partners-indian-government-drive-innovation-through-startups-and-industry-academia",
]
ASSUMPTIONS = [
    ("Consumer share of non-industrial I&I remainder", "40%", "Revenue_Analysis!B18", "Illustrative; no source."),
    ("Industrial share of I&I applied FY2023-25", "42.5% (report point estimate)", "Raw_Data K05", "Only FY2025 given; held constant."),
    ("Industrial GM / OM premium over I&I average", "+2.0pp / 0.0pp", "Profitability_Quality!C24, C26", "Proxy; reconciles to I&I."),
    ("Consumer GM / OM discount vs remainder", "-4.0pp / -2.0pp", "Profitability_Quality!D25, D27", "Proxy; IoT adjusts so weights net to zero."),
    ("Revenue-quality factor weights and scores", "see sheet", "Profitability_Quality!B51:F58", "Analyst assessment."),
    ("Risk factor weights and scores (1-10)", "see sheet", "Risk_Analysis!B11:F21", "Analyst assessment anchored to report."),
    ("Mitigation effectiveness per risk", "30%-50%", "Risk_Analysis!I30:I36", "Analyst assumption."),
    ("Market-opportunity factor scores", "see sheet", "Growth_Opportunity!B22:F26", "Analyst assessment."),
    ("Attach-rate model: baseline attach 1.0; relative ASP 0.30", "1.0 / 0.30", "Growth_Opportunity!B50, B52", "INDEX units; illustrative."),
    ("Scorecard criterion weights (assignment suggestion) and subjective scores", "20/15/15/15/10/10/10/5", "Strategic_Scorecard!B6:F13", "Editable."),
    ("Non-Industrial scenario probabilities, growth and GM deltas", "see sheet", "Forecast_Scenarios!B11:D58", "Illustrative scenarios."),
    ("GM-to-OM pass-through", "80%", "Forecast_Scenarios!B5", "Assumption."),
    ("Incremental commercial budget", "JPY 4.0 bn", "Resource_Allocation!B51", "PLACEHOLDER - replace with approved budget."),
    ("Attributable revenue uplift (% of FY2026E Industrial revenue)", "2.5% / 5.0% / 0.5%", "Resource_Allocation!B54:D54", "Assumption: report does not attribute revenue to the programme."),
    ("Recommended resource allocations (category / application / submarket)", "see sheet", "Resource_Allocation!C6:C11, D29:D32, D38:D44", "Analyst judgement."),
]


def build_sources(ws):
    title(ws, "SOURCES AND ASSUMPTIONS - audit trail", "Every source value links back to Raw_Data. Types: Reported* (as cited in report) | Report estimate | Forecast (report) | Third-party | Calculated | Analyst assumption. Research cutoff: 2 Nov 2026.", 9)
    widths(ws, [8, 56, 14, 12, 14, 26, 24, 18, 60])
    ws.freeze_panes = "A4"
    section(ws, 4, "A. FORMAT LEGEND", 9)
    leg = [("src", "Source input - figure taken from the research report (blue text, light-blue fill)"),
           ("assump", "Analyst assumption / judgement - EDIT THESE to test (blue text, yellow fill)"),
           ("calc", "Formula (black text, no fill) - do not overwrite"),
           ("out", "Key formula output (bold, green fill)"),
           ("grey", "Supporting information / labels (grey fill)"),
           ("flag", "Warning / proxy / flagged item (red)")]
    for i, (k, t) in enumerate(leg):
        W(ws, f"A{5+i}", "", k); W(ws, f"B{5+i}", t, k)
    W(ws, "D5", "Red conditional fill = failed check or reversed lead; green = pass / Industrial leads.", "note", border=False)
    section(ws, 12, "B. Key inputs (live link to Raw_Data) - metric, value, unit, period, source, page, type", 9)
    header(ws, 13, ["ID", "Metric", "Value", "Unit", "Period", "Data type", "Report section / PDF page", "Works-cited ref", "Methodology / note"], h=30)
    n0 = 21
    last_kv = int(R["raw_ids_range"].split("$")[-1])
    for i, rr in enumerate(range(n0, last_kv + 1)):
        r = 14 + i
        W(ws, f"A{r}", f"=Raw_Data!A{rr}", "grey"); W(ws, f"B{r}", f"=Raw_Data!B{rr}", "label", wrap=True)
        W(ws, f"C{r}", f"=Raw_Data!C{rr}", "calc", "#,##0.0##", h="right"); W(ws, f"D{r}", f"=Raw_Data!D{rr}", "grey")
        W(ws, f"E{r}", f"=Raw_Data!E{rr}", "grey"); W(ws, f"F{r}", f"=Raw_Data!F{rr}", "grey", wrap=True)
        W(ws, f"G{r}", f"=Raw_Data!G{rr}", "grey"); W(ws, f"H{r}", f"=Raw_Data!H{rr}", "grey"); W(ws, f"I{r}", f"=Raw_Data!I{rr}", "grey", wrap=True)
        ws.row_dimensions[r].height = 28
    r = 14 + (last_kv - n0 + 1) + 1
    section(ws, r, "C. Derived / calculated metrics and methodology", 9)
    header(ws, r + 1, ["#", "Metric", "Value", "Unit", "Period", "Type", "Location", "", "Methodology"], h=24)
    der = [("Pure Industrial FY2025 (point)", "=Revenue_Analysis!D58", "JPY bn", "FY2025", "Calculated from estimate", "Revenue_Analysis!D58", "I&I reported 671.8 x 42.5% share (report estimate). NOT reported revenue."),
           ("Industrial uncertainty band (high-low)", "=Revenue_Analysis!D60", "JPY bn", "FY2025", "Calculated", "Revenue_Analysis!D60", "(45% - 40%) x I&I."),
           ("Implied Automotive gross margin", f"={R['auto_gm']}", "%", "FY2025", "Derived", "Profitability_Quality!B16", "(Consolidated GM x total - I&I GM x I&I - GM x Other) / Automotive."),
           ("Implied Automotive operating margin", f"={R['auto_om']}", "%", "FY2025", "Derived", "Profitability_Quality!B17", "Same method on operating margin."),
           ("Industrial expected FY2027 revenue", f"={R['ind_exp_rev']}", "JPY bn", "FY2027", "Forecast (probability-weighted)", "Forecast_Scenarios", "SUMPRODUCT(60/20/20, 365/410/295) - report scenarios, not guaranteed."),
           ("Top-ranked application", f"={R['sc_winner']}", "", "FY2027", "Calculated (scorecard)", "Strategic_Scorecard!C30", "Highest SUMPRODUCT(weights, scores)."),
           ("Top Industrial submarket", f"={R['sub_winner']}", "", "FY2027", "Calculated (scorecard)", "Strategic_Scorecard!B60", "Highest weighted score under selected weight set.")]
    for i, (m, f, u, p, t, loc, me) in enumerate(der):
        rr = r + 2 + i
        W(ws, f"A{rr}", f"D{i+1}", "grey"); W(ws, f"B{rr}", m, "label"); W(ws, f"C{rr}", f, "calc", "#,##0.0##" if "margin" not in m else "pct", h="right")
        W(ws, f"D{rr}", u, "grey"); W(ws, f"E{rr}", p, "grey"); W(ws, f"F{rr}", t, "grey", wrap=True); W(ws, f"G{rr}", loc, "grey"); W(ws, f"I{rr}", me, "grey", wrap=True)
        ws.row_dimensions[rr].height = 28
    r = r + 2 + len(der) + 1
    section(ws, r, "D. Analyst assumption register (yellow cells)", 9)
    header(ws, r + 1, ["#", "Assumption", "Value", "", "", "", "Cell location", "", "Comment"], h=24)
    for i, (a, v, loc, cm) in enumerate(ASSUMPTIONS):
        rr = r + 2 + i
        W(ws, f"A{rr}", f"A{i+1}", "grey"); W(ws, f"B{rr}", a, "label", wrap=True); W(ws, f"C{rr}", v, "assump", wrap=True); ws.merge_cells(f"C{rr}:F{rr}")
        W(ws, f"G{rr}", loc, "grey", wrap=True); W(ws, f"I{rr}", cm, "grey", wrap=True); ws.row_dimensions[rr].height = 30
    r = r + 2 + len(ASSUMPTIONS) + 1
    section(ws, r, "E. Research report works cited (reproduced from the report; URLs line-wrap artefacts repaired - verify before reliance)", 9)
    header(ws, r + 1, ["Ref", "Title", "URL", "", "", "", "", "", ""], h=22)
    for i, rf in enumerate(REFS):
        rr = r + 2 + i
        t, u = rf.rsplit("|", 1)
        W(ws, f"A{rr}", f"[{i+1}]", "grey"); W(ws, f"B{rr}", t, "label", wrap=True); W(ws, f"C{rr}", u, "grey")
        ws.row_dimensions[rr].height = 28 if len(t) > 60 else 16
    r = r + 2 + len(REFS) + 1
    section(ws, r, "F. Dates", 9)
    text_rows(ws, r + 1, ["Research cutoff: 2 November 2026 (report). Assignment deadline: 4 November 2026 EOD. Workbook built on 3 November 2026.",
                          "No information after the research cutoff was independently verified; any post-cutoff developments are NOT reflected. Report PDF page numbers are approximate section locations."])


# --------------------------------------------------------------------------------------------- DASHBOARD
def build_dashboard(ws):
    title(ws, "EXECUTIVE DASHBOARD - Where should the semiconductor company focus?",
          "Renesas Electronics case | Research-derived data (no assignment dataset supplied) | Industrial = ESTIMATE | Consumer/IoT = ILLUSTRATIVE | FY2026E = annualised estimate | Research cutoff 2 Nov 2026", 12)
    widths(ws, [24] + [13] * 11)
    ws.sheet_properties.tabColor = NAVY
    # recommendation box
    ws.merge_cells("A4:L4")
    W(ws, "A4", f'="RECOMMENDATION: focus incremental commercial resources on "&{R["sc_winner"]}&" (weighted score "&TEXT(MAX({R["sc_total"]["Automotive"]},{R["sc_total"]["Industrial"]},{R["sc_total"]["Consumer"]},{R["sc_total"]["IoT"]}),"0.00")&", lead "&TEXT({R["sc_lead"]},"0.00")&" pts), and within it on "&{R["sub_winner"]}&" (lead "&TEXT({R["sub_lead"]},"0.00")&" pts). Allocate "&TEXT({R["app_alloc"]["Industrial"]},"0%")&" of incremental resources to Industrial, defend Automotive ("&TEXT({R["app_alloc"]["Automotive"]},"0%")&"). Sensitivity: "&{R["sens_weights"]}&"."', "out", wrap=True, size=11)
    ws.row_dimensions[4].height = 52
    # KPI cards
    cards = [("A", "B", "TOTAL REVENUE FY2025 (A)", f"=Revenue_Analysis!D11", "#,##0.0", f'="JPY bn | FY2026E est. "&TEXT(Revenue_Analysis!E11,"#,##0.0")'),
             ("C", "D", "REVENUE CAGR FY23-25 (actuals)", "=Revenue_Analysis!F11", "pct", '="FY2026E vs FY2025 (est.): "&TEXT(Revenue_Analysis!E36,"+0.0%;-0.0%")'),
             ("E", "F", "HIGHEST EXPECTED GROWTH FY27", "=INDEX($A$17:$A$20,MATCH(MAX($F$17:$F$20),$F$17:$F$20,0))", None, '=TEXT(MAX($F$17:$F$20),"+0.0%")&" (scenario-weighted)"'),
             ("G", "H", "HIGHEST OPERATING MARGIN (derived/proxy)", "=INDEX($A$17:$A$20,MATCH(MAX($H$17:$H$20),$H$17:$H$20,0))", None, '=TEXT(MAX($H$17:$H$20),"0.0%")&" | top GM: "&INDEX($A$17:$A$20,MATCH(MAX($G$17:$G$20),$G$17:$G$20,0))&" "&TEXT(MAX($G$17:$G$20),"0.0%")'),
             ("I", "J", "HIGHEST RISK-ADJUSTED OPPORTUNITY", f"={R['sc_winner']}", None, f'="Score "&TEXT(MAX($J$17:$J$20),"0.00")&" | lead "&TEXT({R["sc_lead"]},"0.00")'),
             ("K", "L", "RECOMMENDED INVESTMENT PRIORITY", f'={R["sc_winner"]}&": "&{R["sub_winner"]}', None, f'="Industrial share of resources: "&TEXT({R["app_alloc"]["Industrial"]},"0%")')]
    for a, b, lab, val, fm, sub in cards:
        ws.merge_cells(f"{a}6:{b}6"); ws.merge_cells(f"{a}7:{b}7"); ws.merge_cells(f"{a}8:{b}8")
        for rr in (6, 7, 8):
            for c in (a, b):
                ws[f"{c}{rr}"].fill = PatternFill("solid", fgColor="EAF0FA" if rr != 7 else "FFFFFF")
        W(ws, f"{a}6", lab, "label", border=False, wrap=True, size=8, bold=True)
        ws[f"{a}6"].font = Font(name=FONT, size=8, bold=True, color=BLUE)
        W(ws, f"{a}7", val, "label", fmt=fm, border=False, h="center", bold=True, size=15, wrap=True)
        ws[f"{a}7"].font = Font(name=FONT, size=15, bold=True, color=NAVY)
        ws[f"{a}7"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        W(ws, f"{a}8", sub, "note", border=False)
        ws[f"{a}8"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[6].height = 26; ws.row_dimensions[7].height = 40; ws.row_dimensions[8].height = 30
    # B14 for validation link: total revenue
    # summary table
    section(ws, 10, "APPLICATION SCORECARD (all figures formula-linked)", 12)
    W(ws, "A11", "Data basis: Automotive = reported; Industrial = estimate (42.5% of I&I); Consumer, IoT = illustrative split; margins = derived/proxy; FY27 growth = probability-weighted scenarios (Industrial from report; others analyst).", "note", border=False)
    W(ws, "A13", "Reference totals", "label", border=False)
    W(ws, "A14", "Total revenue FY2025 (JPY bn)", "label"); W(ws, "B14", "=Revenue_Analysis!D11", "calc", "bn")
    W(ws, "C14", "FY2026E est.", "label", border=False); W(ws, "D14", "=Revenue_Analysis!E11", "calc", "bn")
    W(ws, "E14", "FY2027E expected", "label", border=False); W(ws, "F14", f"={R['rev_total_fc']}", "calc", "bn")
    W(ws, "G14", "growth vs FY26E", "label", border=False); W(ws, "H14", f"={R['rev_total_fc_g']}", "calc", "pct")
    header(ws, 16, ["Application", "Revenue FY2025 (JPY bn)", "Share of total", "CAGR FY23-25 (actuals)", "FY26E YoY (est.)", "FY27E growth (scenario-wtd)", "Gross margin FY25 (proxy)", "Operating margin FY25 (proxy)", "Weighted risk (high = riskier)", "Strategic score (1-10)", "Rank", "Recommended resource share"], h=56)
    for i, app in enumerate(APPS):
        r = 17 + i; rr = R["rev_rows"][app]
        W(ws, f"A{r}", app, "label", bold=True)
        W(ws, f"B{r}", f"=Revenue_Analysis!$D${rr}", "calc", "bn")
        W(ws, f"C{r}", f"=Revenue_Analysis!$D${22+i+1}", "calc", "pct")
        W(ws, f"D{r}", f"=Revenue_Analysis!$F${rr}", "calc", "pct")
        W(ws, f"E{r}", f"=Revenue_Analysis!$E${rr}/Revenue_Analysis!$D${rr}-1", "calc", "pct")
        W(ws, f"F{r}", f"={R['fc_sum_growth'][app]}", "calc", "pct")
        W(ws, f"G{r}", f"={R['gm_proxy'][app]}", "calc", "pct")
        W(ws, f"H{r}", f"={R['om_proxy'][app]}", "calc", "pct")
        W(ws, f"I{r}", f"={R['risk_w'][app]}", "calc", "sc")
        W(ws, f"J{r}", f"={R['sc_total'][app]}", "out", "sc")
        W(ws, f"K{r}", f"={R['sc_rank'][app]}", "out", "int", h="center")
        W(ws, f"L{r}", f"={R['app_alloc'][app]}", "calc", "pct")
    W(ws, "A21", "Total / check", "label", bold=True)
    W(ws, "B21", "=SUM(B17:B20)", "calc", "bn"); W(ws, "C21", "=SUM(C17:C20)", "calc", "pct"); W(ws, "L21", "=SUM(L17:L20)", "out", "pct")
    W(ws, "D21", "(four applications; 'Other' residual excluded)", "note", border=False)
    ws.conditional_formatting.add("D17:F20", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=DRED, bold=True)))
    ws.conditional_formatting.add("D17:F20", CellIsRule(operator="greaterThan", formula=["0.05"], fill=PatternFill("solid", bgColor=GREEN)))
    ws.conditional_formatting.add("I17:I20", ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="F8696B"))
    ws.conditional_formatting.add("J17:J20", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add("K17:K20", CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", bgColor=GREEN), font=Font(color=DGREEN, bold=True)))
    # charts
    S = "Dashboard"
    c1 = BarChart(); c1.type = "col"; c1.title = "1. Revenue by application, FY2025 (A) vs FY2026E (est.), JPY bn"
    rv = wb_ref["Revenue_Analysis"]
    c1.add_data(Reference(rv, min_col=4, max_col=5, min_row=5, max_row=9), titles_from_data=True)
    c1.set_categories(Reference(rv, min_col=1, min_row=6, max_row=9))
    c1.height, c1.width = 7.6, 15.8
    c1.y_axis.title = "JPY bn"; c1.x_axis.delete = False; c1.y_axis.delete = False
    ws.add_chart(c1, "A24")
    c2 = BarChart(); c2.type = "col"; c2.title = "2. Revenue growth: history vs scenario-weighted FY2027"
    for col, nm in ((4, "CAGR FY23-25 (actuals)"), (6, "FY27E growth (scenario-weighted)")):
        s = Series(Reference(ws, min_col=col, min_row=17, max_row=20), title=nm); c2.series.append(s)
    c2.set_categories(Reference(ws, min_col=1, min_row=17, max_row=20))
    c2.y_axis.number_format = "0%"; c2.height, c2.width = 7.6, 15.8
    c2.x_axis.delete = False; c2.y_axis.delete = False
    ws.add_chart(c2, "G24")
    c3 = BarChart(); c3.type = "col"; c3.title = "3. Profitability FY2025: gross vs operating margin (derived / proxy)"
    for col, nm in ((7, "Gross margin"), (8, "Operating margin")):
        c3.series.append(Series(Reference(ws, min_col=col, min_row=17, max_row=20), title=nm))
    c3.set_categories(Reference(ws, min_col=1, min_row=17, max_row=20))
    c3.y_axis.number_format = "0%"; c3.height, c3.width = 7.6, 15.8
    c3.x_axis.delete = False; c3.y_axis.delete = False
    ws.add_chart(c3, "A40")
    c4 = ScatterChart(); c4.title = "4. Risk vs growth positioning (bubble = application)"; c4.style = 13
    c4.x_axis.title = "Weighted risk score (higher = riskier)"; c4.y_axis.title = "FY27E expected growth"
    c4.y_axis.number_format = "0%"
    for i, app in enumerate(APPS):
        r = 17 + i
        s = Series(Reference(ws, min_col=6, min_row=r), Reference(ws, min_col=9, min_row=r))
        s.tx = SeriesLabel(strRef=StrRef(f"Dashboard!$A${r}"))
        s.marker.symbol = "circle"; s.marker.size = 13
        s.graphicalProperties = GraphicalProperties(); s.graphicalProperties.line.noFill = True
        c4.series.append(s)
    c4.x_axis.scaling.min = 4; c4.x_axis.scaling.max = 8
    c4.height, c4.width = 7.6, 15.8; c4.x_axis.delete = False; c4.y_axis.delete = False
    ws.add_chart(c4, "G40")
    c5 = BarChart(); c5.type = "bar"; c5.title = "5. Strategic attractiveness: weighted score (1-10)"
    c5.add_data(Reference(ws, min_col=10, min_row=17, max_row=20), titles_from_data=False)
    c5.set_categories(Reference(ws, min_col=1, min_row=17, max_row=20)); c5.legend = None
    c5.dataLabels = DataLabelList(); c5.dataLabels.showVal = True
    c5.height, c5.width = 7.6, 15.8; c5.x_axis.delete = False; c5.y_axis.delete = False
    ws.add_chart(c5, "A56")
    rs = wb_ref["Resource_Allocation"]
    c6 = DoughnutChart(); c6.title = "6. Recommended resource allocation by category"
    c6.add_data(Reference(rs, min_col=3, min_row=6, max_row=11), titles_from_data=False)
    c6.set_categories(Reference(rs, min_col=1, min_row=6, max_row=11))
    c6.dataLabels = DataLabelList(); c6.dataLabels.showPercent = True
    c6.height, c6.width = 7.6, 15.8
    ws.add_chart(c6, "G56")
    # resource table + scenario table below charts
    section(ws, 72, "RECOMMENDED COMMERCIAL RESOURCE ALLOCATION (incremental)", 12)
    header(ws, 73, ["Category", "Report", "Recommended", "Change"], h=24)
    for i in range(6):
        r = 74 + i
        W(ws, f"A{r}", f"=Resource_Allocation!A{6+i}", "label", wrap=True)
        W(ws, f"B{r}", f"=Resource_Allocation!B{6+i}", "calc", "pct"); W(ws, f"C{r}", f"=Resource_Allocation!C{6+i}", "out", "pct")
        W(ws, f"D{r}", f"=Resource_Allocation!D{6+i}", "calc", "pp"); ws.row_dimensions[r].height = 28
    W(ws, "A80", "Total", "label", bold=True); W(ws, "B80", "=SUM(B74:B79)", "calc", "pct"); W(ws, "C80", "=SUM(C74:C79)", "out", "pct")
    header(ws, 73, ["Industrial FY2027 scenario (report; research-based, not guaranteed)", "Revenue JPY bn", "Gross margin", "Probability"], start_col=6, h=24)
    for i, k in enumerate(("base", "bull", "bear")):
        r = 74 + i
        W(ws, f"F{r}", k.capitalize(), "label", bold=True); W(ws, f"G{r}", f"={R['ind_rev'][k]}", "calc", "bn")
        W(ws, f"H{r}", f"={R['ind_gm'][k]}", "calc", "pct"); W(ws, f"I{r}", f"={R['ind_prob'][k]}", "calc", "pct")
    W(ws, "F77", "Prob-weighted", "label", bold=True); W(ws, "G77", f"={R['ind_exp_rev']}", "out", "bn")
    W(ws, "H77", f"={R['ind_exp_gp']}/{R['ind_exp_rev']}", "out", "pct"); W(ws, "I77", "=SUM(I74:I76)", "out", "pct")
    W(ws, "F78", "FY2026E base (report est.)", "label"); W(ws, "G78", f"={R['ind_base26']}", "calc", "bn")
    W(ws, "F79", "Expected growth vs FY2026E", "label"); W(ws, "G79", f"={R['ind_exp_growth']}", "out", "pct")
    section(ws, 82, "HOW TO READ / CAVEATS", 12)
    cav = ["- Overall answer: Industrial (led by Robotics & Physical AI) offers the best risk-adjusted growth; Automotive is defended, not abandoned; Consumer is de-prioritised.",
           "- The Industrial recommendation is robust to weight changes (Sensitivity s.1-2) but depends on subjective Competitive-advantage / Strategic-alignment scores and on Altium pull-through being real (Sensitivity s.4, s.9).",
           "- Industrial revenue is an ESTIMATE (40-45% of I&I), not reported; FY2026E is an annualised estimate; FY2027 figures are scenarios, not guarantees.",
           "- Automotive has the highest implied operating margin (derived); Industrial's profit advantage is at gross-margin level. See Data_Validation F-04.",
           "- Edit yellow cells to test assumptions; the workbook recalculates automatically (press F9 if calculation is set to manual)."]
    text_rows(ws, 83, cav)
    ws.freeze_panes = "A4"
    R["dash_ok"] = True


wb_ref = {}

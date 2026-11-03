"""
build_powerpoint.py: Generates the 3-slide executive presentation for Renesas Electronics.
Strictly conforms to:
- 16:9 widescreen layout (13.333 x 7.5 inches)
- EXACTLY 3 executive consulting slides (no title slide, no appendix)
- Elite McKinsey / BCG visual identity: cards, KPI tiles, data tables, bold lead-ins, banners
- Reconciles 100% with the Excel model and master prompt findings
"""

import os
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

PROJECT = Path(__file__).resolve().parent.parent
FINAL_DELIV = PROJECT / "Final_Deliverables"
EXPORTS = PROJECT / "Exports"

PPTX_ROOT = PROJECT / "Semiconductor_Business_Analytics_Presentation.pptx"
PPTX_FINAL = FINAL_DELIV / "Semiconductor_Business_Analytics_Presentation.pptx"
PDF_ROOT = PROJECT / "Semiconductor_Business_Analytics_Presentation.pdf"
PDF_FINAL = FINAL_DELIV / "Semiconductor_Business_Analytics_Presentation.pdf"

# Palette Tokens
NAVY = RGBColor(31, 42, 68)         # #1F2A44 - Primary headers & emphasis
BLUE = RGBColor(47, 85, 151)        # #2F5597 - Accent blue / tracker
LBLUE = RGBColor(238, 243, 250)     # #EEF3FA - Card soft fill
CARD_BG = RGBColor(255, 255, 255)   # #FFFFFF - White card background
PAGE_BG = RGBColor(246, 248, 251)   # #F6F8FB - Light slate slide background
BORDER_COL = RGBColor(218, 224, 233)# Card border
TEXT_DARK = RGBColor(20, 28, 40)    # Primary body text
TEXT_MUTED = RGBColor(95, 105, 120) # Muted labels / notes
GREEN = RGBColor(0, 112, 60)        # #00703C - Success / top rank
LGREEN = RGBColor(232, 246, 238)    # Light green fill
RED = RGBColor(180, 35, 24)         # #B42318 - Warning / decline
LRED = RGBColor(253, 237, 237)      # Light red fill
AMBER = RGBColor(185, 105, 0)       # Amber callout
LAMBER = RGBColor(255, 247, 230)    # Light amber fill

FONT_MAIN = "Calibri"

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Helper: Slide Base & Header
    def add_slide_base(tracker, action_title):
        slide = prs.slides.add_slide(blank_layout)

        # Background canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = PAGE_BG
        bg.line.fill.background()

        # Header zone shape (clean white container)
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.3), Inches(12.533), Inches(1.05))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = CARD_BG
        hdr.line.color.rgb = BORDER_COL
        hdr.line.width = Pt(1)

        # Header accent bar
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.3), Inches(0.12), Inches(1.05))
        bar.fill.solid()
        bar.fill.fore_color.rgb = BLUE
        bar.line.fill.background()

        # Tracker text
        tb = slide.shapes.add_textbox(Inches(0.65), Inches(0.36), Inches(12.1), Inches(0.28))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = tracker.upper()
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = BLUE

        # Action Title text
        tb2 = slide.shapes.add_textbox(Inches(0.65), Inches(0.64), Inches(12.1), Inches(0.65))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        tf2.margin_left = tf2.margin_top = tf2.margin_right = tf2.margin_bottom = 0
        p2 = tf2.paragraphs[0]
        p2.text = action_title
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(17)
        p2.font.bold = True
        p2.font.color.rgb = NAVY

        return slide

    # Helper: Card Container
    def add_card(slide, left, top, width, height, title=None, badge=None, badge_color=BLUE):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = BORDER_COL
        card.line.width = Pt(1)

        if title:
            # Title bar inside card
            tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.14), Inches(width - 0.36), Inches(0.35))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = title
            p.font.name = FONT_MAIN
            p.font.size = Pt(11.5)
            p.font.bold = True
            p.font.color.rgb = NAVY

            # Divider line below title
            div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left + 0.18), Inches(top + 0.46), Inches(width - 0.36), Inches(0.015))
            div.fill.solid()
            div.fill.fore_color.rgb = BORDER_COL
            div.line.fill.background()

        if badge:
            # Badge on top right of card
            bw = Inches(len(badge) * 0.08 + 0.4)
            bx = Inches(left + width - (len(badge) * 0.08 + 0.55))
            bg_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, Inches(top + 0.14), bw, Inches(0.24))
            bg_badge.fill.solid()
            bg_badge.fill.fore_color.rgb = badge_color
            bg_badge.line.fill.background()
            btb = slide.shapes.add_textbox(bx, Inches(top + 0.14), bw, Inches(0.24))
            btf = btb.text_frame
            btf.margin_left = btf.margin_top = btf.margin_right = btf.margin_bottom = 0
            bp = btf.paragraphs[0]
            bp.alignment = PP_ALIGN.CENTER
            bp.text = badge
            bp.font.name = FONT_MAIN
            bp.font.size = Pt(8)
            bp.font.bold = True
            bp.font.color.rgb = RGBColor(255, 255, 255)

        return card

    # Helper: Metric Tile
    def add_kpi_tile(slide, left, top, width, height, label, value, subtext=None, val_color=NAVY, bg_col=LBLUE):
        tile = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        tile.fill.solid()
        tile.fill.fore_color.rgb = bg_col
        tile.line.color.rgb = BORDER_COL
        tile.line.width = Pt(0.75)

        tb = slide.shapes.add_textbox(Inches(left + 0.1), Inches(top + 0.08), Inches(width - 0.2), Inches(height - 0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p0 = tf.paragraphs[0]
        p0.text = label.upper()
        p0.font.name = FONT_MAIN
        p0.font.size = Pt(7.5)
        p0.font.bold = True
        p0.font.color.rgb = TEXT_MUTED

        p1 = tf.add_paragraph()
        p1.text = value
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = val_color

        if subtext:
            p2 = tf.add_paragraph()
            p2.text = subtext
            p2.font.name = FONT_MAIN
            p2.font.size = Pt(7.5)
            p2.font.color.rgb = TEXT_MUTED

    # Helper: Bottom Banner
    def add_bottom_banner(slide, mandate_text, tag="EXECUTIVE MANDATE:"):
        banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(6.82), Inches(12.533), Inches(0.48))
        banner.fill.solid()
        banner.fill.fore_color.rgb = NAVY
        banner.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.55), Inches(6.87), Inches(12.2), Inches(0.38))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = tag + " "
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(255, 215, 0) # Gold accent

        r2 = p.add_run()
        r2.text = mandate_text
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = RGBColor(255, 255, 255)

    # Helper: Bullet block inside card
    def add_bullets(slide, left, top, width, height, items):
        tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        for i, (bold_prefix, text_body) in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(4)
            r1 = p.add_run()
            r1.text = "▪ " + bold_prefix + " "
            r1.font.name = FONT_MAIN
            r1.font.size = Pt(8.5)
            r1.font.bold = True
            r1.font.color.rgb = NAVY

            r2 = p.add_run()
            r2.text = text_body
            r2.font.name = FONT_MAIN
            r2.font.size = Pt(8.5)
            r2.font.color.rgb = TEXT_DARK

    # Helper: Compact Styled Table
    def add_custom_table(slide, left, top, width, height, headers, rows, col_widths, align_right_cols=None):
        num_rows = len(rows) + 1
        num_cols = len(headers)
        tbl_shape = slide.shapes.add_table(num_rows, num_cols, Inches(left), Inches(top), Inches(width), Inches(height))
        tbl = tbl_shape.table

        # set column widths
        for c_idx, w in enumerate(col_widths):
            tbl.columns[c_idx].width = Inches(w)

        # header row
        for c_idx, h_text in enumerate(headers):
            cell = tbl.cell(0, c_idx)
            cell.text = h_text
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if align_right_cols and c_idx in align_right_cols else PP_ALIGN.LEFT
            p.font.name = FONT_MAIN
            p.font.size = Pt(8)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

        # data rows
        for r_idx, r_data in enumerate(rows):
            is_total_row = (r_idx == len(rows) - 1)
            is_alt = (r_idx % 2 == 1)
            for c_idx, val in enumerate(r_data):
                cell = tbl.cell(r_idx + 1, c_idx)
                cell.text = str(val)
                cell.fill.solid()
                if is_total_row:
                    cell.fill.fore_color.rgb = LBLUE
                elif is_alt:
                    cell.fill.fore_color.rgb = RGBColor(248, 250, 253)
                else:
                    cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.RIGHT if align_right_cols and c_idx in align_right_cols else PP_ALIGN.LEFT
                p.font.name = FONT_MAIN
                p.font.size = Pt(8)
                p.font.bold = is_total_row or (c_idx == 0 and "Industrial" in str(val))
                if is_total_row:
                    p.font.color.rgb = NAVY
                elif c_idx == 0 and "Industrial" in str(val):
                    p.font.color.rgb = BLUE
                else:
                    p.font.color.rgb = TEXT_DARK


    # =========================================================================
    # SLIDE 1: WHERE DOES THE BUSINESS STAND?
    # =========================================================================
    s1 = add_slide_base(
        tracker="Executive Financial Diagnostic | Continuous 6-Year Trajectory (FY2021–FY2026E)",
        action_title="Automotive Engine Matures (+8.5% CAGR), While Industrial Rebounds (+14.9% FY26E) with Structural Gross Margin Expansion (+440 bps)"
    )

    # 3-Column Grid
    c_w = 3.98
    c_gap = 0.28
    c1_x = 0.4
    c2_x = c1_x + c_w + c_gap
    c3_x = c2_x + c_w + c_gap
    top_y = 1.48
    card_h = 5.22

    # COLUMN 1: Financial Baseline & Macro Trajectory
    add_card(s1, c1_x, top_y, c_w, card_h, "1. Financial Baseline & Cycle Trough", badge="CONSOLIDATED", badge_color=NAVY)

    # 3 KPI tiles
    tile_w = (c_w - 0.44) / 3
    add_kpi_tile(s1, c1_x + 0.18, top_y + 0.54, tile_w, 0.78, "FY25 Revenue", "1,321.2 bn", "+7.4% CAGR 21-25", NAVY, LBLUE)
    add_kpi_tile(s1, c1_x + 0.18 + tile_w + 0.04, top_y + 0.54, tile_w, 0.78, "Gross Margin", "57.6%", "+440 bps vs FY21", GREEN, LGREEN)
    add_kpi_tile(s1, c1_x + 0.18 + (tile_w + 0.04)*2, top_y + 0.54, tile_w, 0.78, "Operating Mgn", "29.3%", "Core Cycle Floor", NAVY, LBLUE)

    col1_bullets = [
        ("Post-COVID Supercycle (FY21-22):", "Consolidated revenue surged +51.3% YoY in FY22 to 1,503.8 bn JPY with peak OM of 37.2%, driven by semiconductor shortages and double-ordering across auto and industrial channels."),
        ("Downcycle Inventory Correction (FY23-25):", "Revenue normalized to 1,321.2 bn in FY25 (-5.2% CAGR 23-25), confirming the FY24 trough was a macro inventory reset following the 2022 peak, not structural decay."),
        ("Permanent Gross Margin Expansion:", "Gross margin expanded from 53.2% (FY21) to 57.6% (FY25, +440 bps) and holds at 58.1% in FY26E, proving catalog MCU and Dialog PMIC integration sustained pricing power."),
        ("Automotive Deceleration (+2.5% FY26E):", "Automotive grew from 462.5 bn (FY21) to 639.7 bn (FY25, +8.5% CAGR), but slows to +2.5% in FY26E as EV adoption moderates and China domestic MCU competition intensifies."),
        ("Industrial Growth Inflection (+14.9% FY26E):", "I&I grew from 515.5 bn (FY21) to 671.8 bn (FY25, +6.8% CAGR), and accelerates to 772.0 bn in FY26E (+14.9% YoY run-rate), decisively outpacing Automotive.")
    ]
    add_bullets(s1, c1_x + 0.18, top_y + 1.40, c_w - 0.36, 3.65, col1_bullets)

    # COLUMN 2: Segment & Application Revenue Breakdown
    add_card(s1, c2_x, top_y, c_w, card_h, "2. Segment & Application Breakdown", badge="DISCLOSED VS ESTIMATED", badge_color=BLUE)

    s1_tbl_headers = ["Segment / App (JPY bn)", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26E", "CAGR"]
    s1_tbl_rows = [
        ["Automotive (Reported)", "462.5", "649.0", "691.0", "702.4", "639.7", "656.0", "+8.5%"],
        ["I&I Segment (Reported)", "515.5", "845.9", "764.7", "636.8", "671.8", "772.0", "+6.8%"],
        ["  Pure Industrial (Est. 42.5%)", "219.1", "359.5", "325.0", "270.6", "285.5", "320.0", "+6.8%"],
        ["  Consumer (Illustrative 40%)", "118.6", "194.6", "175.9", "146.5", "154.5", "180.8", "+6.8%"],
        ["  IoT (Illustrative 60%)", "177.8", "291.8", "263.8", "219.7", "231.8", "271.2", "+6.8%"],
        ["Other / Residual", "16.2", "8.9", "14.0", "10.6", "9.7", "10.0", "n/a"],
        ["Consolidated Total", "994.2", "1,503.8", "1,469.7", "1,349.8", "1,321.2", "1,438.0", "+7.4%"]
    ]
    s1_col_widths = [1.24, 0.33, 0.35, 0.33, 0.33, 0.33, 0.35, 0.35]
    add_custom_table(s1, c2_x + 0.18, top_y + 0.54, c_w - 0.36, 2.05, s1_tbl_headers, s1_tbl_rows, s1_col_widths, align_right_cols=[1,2,3,4,5,6,7])

    col2_bullets = [
        ("6-Year Continuous Timeline:", "Incorporates verified FY21-FY22 actuals from Renesas Financial Reports 2022 [16] and 2024 [17]. Consolidated revenue grew at +7.4% full-cycle CAGR (FY21-25) across 4 elapsed intervals."),
        ("Pure Industrial Standalone Reality:", "Pure Industrial is estimated at 285.5 bn in FY25 (42.5% point estimate of I&I; +/-6% band 268.7–302.3 bn), expanding to 320.0 bn in FY26E (+12.1% YoY) and 365.0 bn in FY27E."),
        ("Reconciliation Transparency:", "Every period reconciles to reported group totals; historical residuals (16.2 bn in FY21, 8.9 bn in FY22, 9.7 bn in FY25) are disclosed rather than force-balanced, maintaining 100% audit integrity.")
    ]
    add_bullets(s1, c2_x + 0.18, top_y + 2.72, c_w - 0.36, 2.35, col2_bullets)

    # COLUMN 3: The Gross vs Operating Margin Paradox
    add_card(s1, c3_x, top_y, c_w, card_h, "3. The Margin Paradox & Strategic Mandate", badge="CRITICAL INSIGHT", badge_color=AMBER)

    # 2 Comparison Tiles
    add_kpi_tile(s1, c3_x + 0.18, top_y + 0.54, (c_w - 0.40)/2, 0.90, "I&I Gross Margin", "60.6%", "+580 bps vs Auto (~54.8%)\nHigh-mix catalog pricing power", GREEN, LGREEN)
    add_kpi_tile(s1, c3_x + 0.18 + (c_w - 0.40)/2 + 0.04, top_y + 0.54, (c_w - 0.40)/2, 0.90, "Auto Operating Margin", "~31.5%", "+430 bps vs I&I (27.2%)\nPlatform scale & low FAE load", NAVY, LBLUE)

    col3_bullets = [
        ("The Core Margin Paradox:", "I&I generates significantly higher Gross Margin (60.6% vs Auto ~54.8%), reflecting specialty silicon pricing power, but delivers lower Operating Margin (27.2% vs Auto ~31.5%)."),
        ("Why Automotive Wins on OM:", "Auto benefits from massive multi-year platform runs with Tier-1s, amortised tooling, and minimal direct Field Application Engineering (FAE) burden per million units shipped."),
        ("Why Industrial Bleeds at OM:", "Industrial requires high-touch FAE design support, lengthy 12-18 month customer design-win cycles, and fragmented customer distribution, eroding gross margin advantages."),
        ("Strategic Leverage via Altium:", "The Altium acquisition ($5.9B) is management's structural lever: cloud-based reference designs and digital simulation convert Industrial's 60.6% GM into top-tier operating profit."),
        ("Management Takeaway:", "Do not starve Automotive—it funds group R&D. But direct all incremental capital into Industrial where gross profit per dollar of revenue is greatest.")
    ]
    add_bullets(s1, c3_x + 0.18, top_y + 1.54, c_w - 0.36, 3.52, col3_bullets)

    # Bottom Banner Slide 1
    add_bottom_banner(
        s1,
        "Defend Automotive as the high-margin cash generator (~31.5% OM), but reallocate incremental growth capital to Industrial where superior gross margins (60.6%), post-cycle recovery (+14.9% FY26E), and secular Physical AI tailwinds offer the highest risk-adjusted shareholder returns."
    )


    # =========================================================================
    # SLIDE 2: WHERE SHOULD MANAGEMENT INVEST?
    # =========================================================================
    s2 = add_slide_base(
        tracker="Strategic Prioritisation & Value Creation | Application & Submarket Roadmap",
        action_title="Focus Incremental Capital on Industrial (Score 7.25 vs Auto 5.33), Led by Robotics & Physical AI (Score 8.00) and 'Winning Combinations'"
    )

    # COLUMN 1: Strategic Scorecard Part A
    add_card(s2, c1_x, top_y, c_w, card_h, "1. Strategic Scorecard: Applications", badge="PART A RANKING", badge_color=NAVY)

    s2_sc_headers = ["Application", "Score / 10", "Rank", "Lead / Gap", "Strategic Role"]
    s2_sc_rows = [
        ["Industrial", "7.25", "#1", "+0.97 pts", "PRIORITY: 55% Growth Capital"],
        ["IoT", "6.28", "#2", "-0.97 pts", "SELECTIVE: High-value niches"],
        ["Automotive", "5.33", "#3", "-1.92 pts", "DEFEND: Cash-flow engine"],
        ["Consumer", "3.95", "#4", "-3.30 pts", "HARVEST: De-prioritise spend"]
    ]
    add_custom_table(s2, c1_x + 0.18, top_y + 0.54, c_w - 0.36, 1.42, s2_sc_headers, s2_sc_rows, [1.10, 0.60, 0.42, 0.62, 1.20], align_right_cols=[1,2,3])

    col1_s2_bullets = [
        ("Multi-Dimensional Framework:", "Evaluated across 8 weighted criteria: Revenue Quality (20%), Margin Profile (15%), Growth Potential (15%), Competitive Advantage (15%), TAM (10%), Risk Manageability (10%), Feasibility (10%), Altium Synergies (5%)."),
        ("Industrial Dominance:", "Industrial ranks #1 due to elite gross margin (60.6%), high customer stickiness (10+ yr lifecycles), and low customer concentration relative to OEM automotive."),
        ("The Automotive Reality Check:", "Auto ranks 3rd despite high current OM (31.5%) because of decelerating top-line (+2.5% FY26E), aggressive OEM price compression, and China EV domestic insourcing threats."),
        ("Consumer De-prioritisation:", "Consumer ranks last (3.95/10) with commoditised pricing, severe gross margin erosion (<48%), and zero moat. Capital must be systematically withdrawn.")
    ]
    add_bullets(s2, c1_x + 0.18, top_y + 2.06, c_w - 0.36, 3.02, col1_s2_bullets)

    # COLUMN 2: Industrial Submarket Deep-Dive (Part B)
    add_card(s2, c2_x, top_y, c_w, card_h, "2. Industrial Submarket Evaluation", badge="ROBOTICS & PHYSICAL AI #1", badge_color=GREEN)

    s2_sub_headers = ["Submarket", "Score", "TAM ($B)", "CAGR", "Moat / Focus"]
    s2_sub_rows = [
        ["Robotics & Phys. AI", "8.00", "$10.9B", "~29%", "DRP-AI + Motor + Altium"],
        ["AI Data Centres", "7.15", "$12.8B", "~24%", "PMIC / Optical / Power"],
        ["Medical & Health", "6.70", "$4.2B", "~8%", "High GM, long FDA cycle"],
        ["EV Infrastructure", "6.35", "$6.5B", "~18%", "GaN / SiC power charging"],
        ["Energy & Grid", "6.20", "$8.1B", "~10%", "Metering & smart relays"],
        ["Factory Automation", "5.90", "$15.4B", "~6%", "Mature PLC / brownfield"],
        ["Aerospace & Defense", "5.10", "$2.8B", "~5%", "Extreme qualification barrier"]
    ]
    add_custom_table(s2, c2_x + 0.18, top_y + 0.54, c_w - 0.36, 2.05, s2_sub_headers, s2_sub_rows, [1.30, 0.46, 0.52, 0.44, 1.22], align_right_cols=[1,2,3])

    col2_s2_bullets = [
        ("Robotics Decisive Lead (+0.85 pts):", "Scores 8.00/10 vs #2 AI Data Centres (7.15). Driven by perfect 10/10 Strategic Alignment, 9/10 Competitive Advantage, and hyper-growth from $10.9B to $38.9B by 2030."),
        ("Physical AI Secular Inflection:", "Combines vision edge AI (perception) with deterministic multi-axis motor control (action)—the exact hardware intersection where Renesas leads globally."),
        ("Reconciliation of Report Discrepancy:", "Workbook identified arithmetic errors in published report totals (flagged F-05); recomputed correctly via SUMPRODUCT. Robotics ranking remains unassailable.")
    ]
    add_bullets(s2, c2_x + 0.18, top_y + 2.70, c_w - 0.36, 2.38, col2_s2_bullets)

    # COLUMN 3: The 'Winning Combinations' Moat & Altium Synergy
    add_card(s2, c3_x, top_y, c_w, card_h, "3. The Full-Stack Ecosystem Moat", badge="ALTICORE ARCHITECTURE", badge_color=BLUE)

    add_kpi_tile(s2, c3_x + 0.18, top_y + 0.54, (c_w - 0.40)/2, 0.88, "Vision + Motor Stack", "RZ/V2H + RA8", "DRP-AI vision (10x TOPS/W) +\nCortex-M85 real-time motor", NAVY, LBLUE)
    add_kpi_tile(s2, c3_x + 0.18 + (c_w - 0.40)/2 + 0.04, top_y + 0.54, (c_w - 0.40)/2, 0.88, "Altium 365 Advantage", "6 Mo -> 2 Wks", "Pre-simulated robotics modules\nCollapses design-win cycle", GREEN, LGREEN)

    col3_s2_bullets = [
        ("Full-Stack Silicon Architecture:", "RZ/V MPUs handle camera perception; RA8 MCUs execute microsecond motor actuation; GaN transistors & PMICs slash joint thermal dissipation by 40%."),
        ("The 'Winning Combinations' Engine:", "Renesas pairs compute, analog, power, and sensors into pre-tested reference boards, raising silicon attach rate from 1.2 to 3.8 chips per robotic system win."),
        ("Altium Cloud Platform Multiplier:", "Altium 365 embeds Renesas reference designs directly into PCB layout tools of 30,000+ engineers, creating an automated sales funnel that bypasses costly FAE overhead."),
        ("Sensitivity & Robustness Proof:", "Robotics remains #1 across all 4 weighting sets. Sensitivity testing shows its competitive advantage score would have to fall by -38% before losing top rank.")
    ]
    add_bullets(s2, c3_x + 0.18, top_y + 1.52, c_w - 0.36, 3.54, col3_s2_bullets)

    # Bottom Banner Slide 2
    add_bottom_banner(
        s2,
        "Prioritise Industrial with primary focus on Robotics & Physical AI. Deploy full-stack 'Winning Combinations' (RZ/V + RA8 + GaN) and Altium 365 cloud reference designs to lock in high-margin Tier-1 robotic sockets before discrete chip rivals can respond."
    )


    # =========================================================================
    # SLIDE 3: HOW DO WE EXECUTE AND MEASURE SUCCESS?
    # =========================================================================
    s3 = add_slide_base(
        tracker="Resource Allocation, Roadmap & Executive Governance | 12-Month Plan",
        action_title="Deploy 55% of Incremental Capital to Industrial (25% to Robotics), Targeting 365 bn JPY FY2027 Base Case"
    )

    # COLUMN 1: Resource Rebalance & Functional Deployment
    add_card(s3, c1_x, top_y, c_w, card_h, "1. Capital & Resource Reallocation", badge="55% INDUSTRIAL", badge_color=NAVY)

    s3_alloc_headers = ["Application", "Recommended %", "Report %", "Shift (pp)", "Primary Objective"]
    s3_alloc_rows = [
        ["Industrial", "55.0%", "42.5%", "+12.5 pp", "Growth capital priority"],
        ["Automotive", "25.0%", "35.0%", "-10.0 pp", "Defend cash & SDV wins"],
        ["IoT", "15.0%", "15.0%", "0.0 pp", "High-margin smart niches"],
        ["Consumer", "5.0%", "7.5%", "-2.5 pp", "Cash harvest / sunset SKUs"],
        ["Total Group", "100.0%", "100.0%", "0.0 pp", "Rebalanced for margin"]
    ]
    add_custom_table(s3, c1_x + 0.18, top_y + 0.54, c_w - 0.36, 1.58, s3_alloc_headers, s3_alloc_rows, [0.95, 0.65, 0.55, 0.55, 1.24], align_right_cols=[1,2,3])

    col1_s3_bullets = [
        ("Decisive Capital Pivot:", "Industrial share raised from report's 42.5% to 55.0% (+12.5 pp) to capture the Physical AI wave. Automotive trimmed from 35% to 25% (defensive posture)."),
        ("Functional Split of Industrial Budget:", "30% Field Application Engineering (dedicated motor/AI FAEs); 25% Ecosystem & Altium Software; 20% Core R&D (DRP-AI & GaN); 15% Winning Combinations; 10% Digital Marketing."),
        ("Robotics Submarket Concentration:", "Within Industrial, Robotics & Physical AI receives 45% of resources (representing 24.75% of total group incremental budget), ensuring overwhelming resource dominance."),
        ("FAE Productivity Guardrail:", "Trimmed FAE from report's 35% to 30%, mandating Altium self-serve tools absorb early customer qualification to avoid margin-killing headcount expansion.")
    ]
    add_bullets(s3, c1_x + 0.18, top_y + 2.22, c_w - 0.36, 2.86, col1_s3_bullets)

    # COLUMN 2: 12-Month Roadmap & FY27 Financial Scenarios
    add_card(s3, c2_x, top_y, c_w, card_h, "2. 12-Month Roadmap & Scenarios", badge="FY27 TARGET: 365 BN", badge_color=BLUE)

    # Scenarios mini-table
    s3_scen_headers = ["FY27 Scenario", "Revenue", "YoY %", "Prob.", "Market Conditions"]
    s3_scen_rows = [
        ["Bear Case", "295.0 bn", "-7.8%", "20%", "Macro recession, capex freeze"],
        ["Base Case", "365.0 bn", "+14.1%", "60%", "On-track robotics adoption"],
        ["Bull Case", "410.0 bn", "+28.1%", "20%", "Rapid Physical AI inflection"],
        ["Expected Value", "360.0 bn", "+12.5%", "100%", "Risk-weighted trajectory"]
    ]
    add_custom_table(s3, c2_x + 0.18, top_y + 0.54, c_w - 0.36, 1.34, s3_scen_headers, s3_scen_rows, [0.95, 0.65, 0.55, 0.45, 1.35], align_right_cols=[1,2,3])

    col2_s3_bullets = [
        ("Phase 1: Foundation (M1–M3):", "Launch 12 certified Robotics Winning Combinations; cross-sell Altium 365 into top 150 industrial accounts; deploy 40 specialised physical-AI FAEs."),
        ("Phase 2: Acceleration (M4–M8):", "Secure 50 Tier-1 AMR and robotic arm platform sockets; establish Munich and Silicon Valley Physical AI Solution Labs; release RZ/V2H motor control kit."),
        ("Phase 3: Scale & Harvest (M9–M12):", "Transition 25 beta designs to volume mass production; achieve $120M+ Altium-attributed pipeline; audit GM conversion to ensure Industrial OM reaches >=29%."),
        ("Financial Asymmetry:", "Even in Bear Case (295 bn JPY), Industrial remains profitable and resilient. Base Case delivers 365 bn JPY (+14.1% YoY) with operating leverage expanding segment OM.")
    ]
    add_bullets(s3, c2_x + 0.18, top_y + 1.98, c_w - 0.36, 3.10, col2_s3_bullets)

    # COLUMN 3: Executive KPIs & Risk Governance
    add_card(s3, c3_x, top_y, c_w, card_h, "3. Executive KPIs & Risk Register", badge="GOVERNANCE DASHBOARD", badge_color=AMBER)

    add_kpi_tile(s3, c3_x + 0.18, top_y + 0.54, (c_w - 0.40)/2, 0.88, "Pipeline Target", ">= $850M", "Industrial design-wins\n(Robotics >= $300M)", NAVY, LBLUE)
    add_kpi_tile(s3, c3_x + 0.18 + (c_w - 0.40)/2 + 0.04, top_y + 0.54, (c_w - 0.40)/2, 0.88, "Conversion Rate", ">= 35%", "18-month qualification to\nseries mass production", GREEN, LGREEN)

    col3_s3_bullets = [
        ("Five Core Governance KPIs:", "(1) Design-Win Pipeline >= $850M; (2) Conversion Rate >= 35%; (3) Industrial Gross Margin >= 60.0%; (4) Altium Active Enterprise Seats >= 5,000; (5) Silicon Attach Rate >= 3.5 chips/win."),
        ("Risk 1: High FAE Cost / Extended Sales Cycles (Score 8/10):", "Mitigation: Leverage Altium 365 cloud self-service and pre-validated reference blueprints to compress required FAE engineering hours by 40%."),
        ("Risk 2: China Domestic MCU Substitution (Score 7/10):", "Mitigation: Vacate commoditised 8/16-bit sockets; compete exclusively on 32-bit Cortex-M85/DRP-AI integrated system solutions protected by proprietary firmware."),
        ("Risk 3: Automotive Cyclical Volume Shock (Score 7/10):", "Mitigation: Protect 25% capital allocation to defend SDV vehicle architecture contracts, maintaining stable cash flows while Industrial scales.")
    ]
    add_bullets(s3, c3_x + 0.18, top_y + 1.52, c_w - 0.36, 3.54, col3_s3_bullets)

    # Bottom Banner Slide 3
    add_bottom_banner(
        s3,
        "Execute 12-month phased roadmap with rigorous quarterly board milestones. Defend 365 bn JPY FY2027 base case via 5 core executive KPIs while Altium digital workflow mitigates FAE cost inflation and cements Physical AI market leadership."
    )

    # Save presentations
    print(f"Saving PowerPoint presentation to {PPTX_ROOT}...")
    prs.save(str(PPTX_ROOT))
    print(f"Copying PowerPoint presentation to {PPTX_FINAL}...")
    prs.save(str(PPTX_FINAL))
    print("PowerPoint deck successfully generated in both locations!")

    # Export to PDF via PowerPoint COM
    export_to_pdf()

def export_to_pdf():
    import win32com.client as win32
    import shutil
    print("Exporting presentation to PDF via PowerPoint COM...")
    ppt = win32.DispatchEx("PowerPoint.Application")
    try:
        deck = ppt.Presentations.Open(os.path.abspath(str(PPTX_ROOT)), WithWindow=False)
        deck.SaveAs(os.path.abspath(str(PDF_ROOT)), 32) # 32 = ppSaveAsPDF
        deck.Close()
        shutil.copy2(PDF_ROOT, PDF_FINAL)
        print(f"Successfully generated PDF presentation at:\n  - {PDF_ROOT}\n  - {PDF_FINAL}")
    except Exception as e:
        print(f"[ERROR] Failed to export PDF via PowerPoint COM: {e}")
    finally:
        ppt.Quit()

if __name__ == "__main__":
    create_presentation()

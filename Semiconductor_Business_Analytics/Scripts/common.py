"""Shared helpers: paths, colours, cell styling. No machine-specific paths (everything is relative)."""
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

PROJECT = Path(__file__).resolve().parent.parent
FINAL = PROJECT / "Final_Deliverables"
EXPORTS = PROJECT / "Exports"
XLSX_NAME = "Semiconductor_Business_Analytics.xlsx"
PPTX_NAME = "Semiconductor_Business_Analytics_Presentation.pptx"
PDF_README = "README_Setup_and_Methodology.pdf"
PDF_SLIDES = "Semiconductor_Business_Analytics_Presentation.pdf"
RAW_BUILD = PROJECT / "Exports" / "_build_raw.xlsx"   # pre-recalculation workbook

NAVY, BLUE, LBLUE, GREY, YELLOW, GREEN, RED = "1F2A44", "2F5597", "DDEBF7", "F2F2F2", "FFF2CC", "C6EFCE", "FFC7CE"
DGREEN, DRED = "00703C", "B42318"
FONT = "Calibri"

APPS = ["Automotive", "Industrial", "Consumer", "IoT"]
YEARS = ["FY2021 (A)", "FY2022 (A)", "FY2023 (A)", "FY2024 (A)", "FY2025 (A)", "FY2026E (est.)"]

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

FMT = {"bn": '#,##0.0;[Red]-#,##0.0', "pct": '0.0%;[Red]-0.0%', "sc": '0.00', "int": '0', "pp": '+0.0%;[Red]-0.0%;0.0%', "x": '0.00"x"'}

KIND = {
    "src":    dict(font=Font(name=FONT, size=10, color="0000FF"), fill=PatternFill("solid", fgColor=LBLUE)),
    "assump": dict(font=Font(name=FONT, size=10, color="0000FF"), fill=PatternFill("solid", fgColor=YELLOW)),
    "calc":   dict(font=Font(name=FONT, size=10, color="000000"), fill=None),
    "out":    dict(font=Font(name=FONT, size=10, color="000000", bold=True), fill=PatternFill("solid", fgColor=GREEN)),
    "label":  dict(font=Font(name=FONT, size=10, color="000000"), fill=None),
    "note":   dict(font=Font(name=FONT, size=9, color="595959", italic=True), fill=None),
    "grey":   dict(font=Font(name=FONT, size=10, color="404040"), fill=PatternFill("solid", fgColor=GREY)),
    "flag":   dict(font=Font(name=FONT, size=10, color=DRED, bold=True), fill=PatternFill("solid", fgColor=RED)),
}


def W(ws, addr, value, kind="calc", fmt=None, bold=None, wrap=False, h=None, border=True, size=None):
    c = ws[addr]
    c.value = value
    k = KIND[kind]
    f = k["font"]
    if bold is not None or size:
        f = Font(name=f.name, size=size or f.size, color=f.color, bold=f.bold if bold is None else bold, italic=f.italic)
    c.font = f
    if k["fill"] is not None:
        c.fill = k["fill"]
    if fmt:
        c.number_format = FMT.get(fmt, fmt)
    c.alignment = Alignment(wrap_text=wrap, vertical="top" if wrap else "center", horizontal=h)
    if border and kind not in ("note",):
        c.border = BORDER
    return c


def title(ws, text, sub=None, ncols=10):
    for i in range(1, ncols + 1):
        ws.cell(row=1, column=i).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(row=2, column=i).fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].value = text
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color="FFFFFF")
    ws["A2"].value = sub or ""
    ws["A2"].font = Font(name=FONT, size=10, italic=True, color="D9E2F3")
    ws.row_dimensions[1].height = 26
    ws.sheet_view.showGridLines = False


def section(ws, row, text, ncols=10):
    for i in range(1, ncols + 1):
        ws.cell(row=row, column=i).fill = PatternFill("solid", fgColor=BLUE)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    ws.row_dimensions[row].height = 20


def header(ws, row, values, start_col=1, h=30):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDER
    ws.row_dimensions[row].height = h


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[CL(i)].width = w


def text_rows(ws, row, lines, col=1, width_chars=150, kind="label", bold=False):
    """Write wrapped paragraphs across a wide area using one cell per line (no merges)."""
    for ln in lines:
        c = ws.cell(row=row, column=col, value=ln)
        c.font = Font(name=FONT, size=10, bold=bold, color="262626")
        c.alignment = Alignment(wrap_text=False, vertical="center")
        row += 1
    return row


def A(col, row, absolute=True):
    return f"${CL(col)}${row}" if absolute else f"{CL(col)}{row}"

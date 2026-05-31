"""
payslip_core.py
===============
Core pay slip generation logic — no GUI dependencies.
Imported by app.py (GUI) and generate_payslips.py (CLI).
"""

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import os

# ─── Column indices (0-based, matching Excel sheet) ───────────────────────────
COL_EMP_NO        = 0
COL_NAME          = 1
COL_DEPARTMENT    = 2
COL_DESIGNATION   = 3
COL_BASIC         = 4
COL_ALLOWANCE     = 5
COL_ATT_BONUS     = 6
COL_FIXED_ALLOW   = 7
COL_INCENTIVE     = 8
COL_NORMAL_OT     = 9
COL_NORMAL_OT_AMT = 10
COL_DOUBLE_OT     = 11
COL_DOUBLE_OT_AMT = 12
COL_INCENTIVE2    = 13
COL_GROSS         = 14
COL_SAL_ADV       = 15
COL_NO_PAY        = 16
COL_NO_PAY_AMT    = 17
COL_ATT_BON_DED   = 18
COL_ALLOW_DED     = 19
COL_EPF8          = 20
COL_LATE          = 21
COL_LATE_DED      = 22
COL_WELFARE       = 23
COL_TOTAL_DED     = 24
COL_NET_SALARY    = 25
COL_EPF12         = 26
COL_ETF3          = 27

# ─── Slip dimensions ──────────────────────────────────────────────────────────
SLIP_WIDTH  = 5.0 * cm
SLIP_HEIGHT = 12.0 * cm
GAP         = 2 * mm

MARGIN_LEFT   = 0.7 * cm
MARGIN_RIGHT  = 0.7 * cm
MARGIN_TOP    = 0.7 * cm
MARGIN_BOTTOM = 0.7 * cm

# ─── Colours ──────────────────────────────────────────────────────────────────
C_HDR_BG      = colors.HexColor("#1a2744")
C_HDR_TEXT    = colors.white
C_BORDER      = colors.HexColor("#2d3a6b")
C_BORDER_LITE = colors.HexColor("#c8d0e8")
C_ROW_ALT     = colors.HexColor("#f4f7fc")
C_GROSS_BG    = colors.HexColor("#fff8e1")
C_DED_BG      = colors.HexColor("#fde8e8")
C_NET_BG      = colors.HexColor("#e8f8ee")
C_EPF_BG      = colors.HexColor("#f0f0f0")
C_TEXT        = colors.HexColor("#12192b")
C_MUTED       = colors.HexColor("#5a6080")
C_ACCENT      = colors.HexColor("#2d3a8c")


# ─── Helpers ──────────────────────────────────────────────────────────────────

def read_employees(excel_path: str) -> list:
    """Return list of rows (each a list of cell values), skipping header row."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    employees = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        employees.append(list(row))
    wb.close()
    return employees


def _g(row: list, col: int):
    try:
        return row[col]
    except IndexError:
        return None


def _fmt(val) -> str:
    if val is None:
        return "0.00"
    try:
        return f"{float(val):,.2f}"
    except (ValueError, TypeError):
        return "0.00"


def _fmt_units(val) -> str:
    if val is None or val == 0:
        return ""
    try:
        v = float(val)
        return f"{v:.2f}".rstrip("0").rstrip(".")
    except (ValueError, TypeError):
        return ""


# ─── Draw one pay slip ────────────────────────────────────────────────────────

def draw_payslip(c: canvas.Canvas, x: float, y: float,
                 row: list, company_name: str, pay_period: str):
    W = SLIP_WIDTH
    H = SLIP_HEIGHT

    # Outer border
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.7)
    c.rect(x, y, W, H, stroke=1, fill=0)

    # Header block
    hdr_h = 1.62 * cm
    c.setFillColor(C_HDR_BG)
    c.rect(x, y + H - hdr_h, W, hdr_h, stroke=0, fill=1)

    c.setFillColor(C_HDR_TEXT)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString(x + W / 2, y + H - 0.38 * cm, company_name)

    c.setFont("Helvetica-Bold", 6)
    c.drawCentredString(x + W / 2, y + H - 0.64 * cm, "Pay Sheet")
    c.drawCentredString(x + W / 2, y + H - 0.88 * cm, pay_period)

    c.setFont("Helvetica", 5.2)
    emp_no = _g(row, COL_EMP_NO)
    emp_no_str = str(int(emp_no)) if emp_no else "-"
    c.drawString(x + 0.12 * cm, y + H - 1.13 * cm, f"Emp No   {emp_no_str}")

    c.setFont("Helvetica-Bold", 5.5)
    name = _g(row, COL_NAME) or ""
    c.drawCentredString(x + W / 2, y + H - 1.38 * cm, name)

    c.setFont("Helvetica", 4.8)
    dept  = _g(row, COL_DEPARTMENT)  or ""
    desig = _g(row, COL_DESIGNATION) or ""
    c.drawString(x + 0.10 * cm, y + H - 1.58 * cm, f"Dept-{dept}   Desig-{desig}")

    # Body rows
    data_rows = [
        ("Basic",            "",                                   _fmt(_g(row, COL_BASIC)),        "earn"),
        ("Allowance",        "",                                   _fmt(_g(row, COL_ALLOWANCE)),    "earn"),
        ("Attendance Bonus", "",                                   _fmt(_g(row, COL_ATT_BONUS)),    "earn"),
        ("Fixed Allowance",  "",                                   _fmt(_g(row, COL_FIXED_ALLOW)),  "earn"),
        ("Incentive",        "",                                   _fmt(_g(row, COL_INCENTIVE)),    "earn"),
        ("Normal OT",        _fmt_units(_g(row, COL_NORMAL_OT)),  _fmt(_g(row, COL_NORMAL_OT_AMT)),"earn"),
        ("Double OT",        _fmt_units(_g(row, COL_DOUBLE_OT)),  _fmt(_g(row, COL_DOUBLE_OT_AMT)),"earn"),
        ("Incentive",        "",                                   _fmt(_g(row, COL_INCENTIVE2)),   "earn"),
        ("Gross Salary",     "",                                   _fmt(_g(row, COL_GROSS)),        "gross"),
        ("Salary Advance",   "",                                   _fmt(_g(row, COL_SAL_ADV)),      "deduct"),
        ("No Pay",           _fmt_units(_g(row, COL_NO_PAY)),     _fmt(_g(row, COL_NO_PAY_AMT)),   "deduct"),
        ("Attendance Bonus", "",                                   _fmt(_g(row, COL_ATT_BON_DED)),  "deduct"),
        ("Allowance",        "",                                   _fmt(_g(row, COL_ALLOW_DED)),    "deduct"),
        ("E.P.F 8%",         "",                                   _fmt(_g(row, COL_EPF8)),         "deduct"),
        ("Late",             _fmt_units(_g(row, COL_LATE)),       _fmt(_g(row, COL_LATE_DED)),     "deduct"),
        ("Welfare",          "",                                   _fmt(_g(row, COL_WELFARE)),      "deduct"),
        ("Total Deduction",  "",                                   _fmt(_g(row, COL_TOTAL_DED)),    "total_d"),
        ("Net Salary",       "",                                   _fmt(_g(row, COL_NET_SALARY)),   "net"),
        ("E.P.F. 12%",       "",                                   _fmt(_g(row, COL_EPF12)),        "epf"),
        ("E.T.F. 3%",        "",                                   _fmt(_g(row, COL_ETF3)),         "epf"),
    ]

    BODY_TOP = y + H - hdr_h
    BODY_H   = BODY_TOP - y
    ROW_H    = BODY_H / len(data_rows)
    COL1_W   = 2.30 * cm
    COL2_W   = 0.68 * cm
    LABEL_X  = x + 0.12 * cm
    UNITS_X  = x + COL1_W + COL2_W / 2
    AMT_X    = x + W - 0.10 * cm

    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.35)
    c.line(x, BODY_TOP, x + W, BODY_TOP)

    for i, (label, units, amount, style) in enumerate(data_rows):
        ry = BODY_TOP - (i + 1) * ROW_H
        if style == "gross":
            c.setFillColor(C_GROSS_BG)
            c.rect(x + 0.05, ry + 0.05, W - 0.1, ROW_H - 0.05, stroke=0, fill=1)
        elif style == "total_d":
            c.setFillColor(C_DED_BG)
            c.rect(x + 0.05, ry + 0.05, W - 0.1, ROW_H - 0.05, stroke=0, fill=1)
        elif style == "net":
            c.setFillColor(C_NET_BG)
            c.rect(x + 0.05, ry + 0.05, W - 0.1, ROW_H - 0.05, stroke=0, fill=1)
        elif style == "epf":
            c.setFillColor(C_EPF_BG)
            c.rect(x + 0.05, ry + 0.05, W - 0.1, ROW_H - 0.05, stroke=0, fill=1)
        elif i % 2 == 0:
            c.setFillColor(C_ROW_ALT)
            c.rect(x + 0.05, ry + 0.03, W - 0.1, ROW_H - 0.03, stroke=0, fill=1)

        c.setStrokeColor(C_BORDER_LITE)
        c.setLineWidth(0.18)
        c.line(x, ry, x + W, ry)

        ty   = ry + (ROW_H * 0.22)
        bold  = style in ("gross", "total_d", "net")
        small = style == "epf"
        label_size = 5.4 if small else 6.2 if bold else 5.7
        amount_size = 5.4 if small else 6.2 if bold else 5.7

        c.setFont("Helvetica-Bold" if bold else "Helvetica", label_size)
        c.setFillColor(C_MUTED if small else C_ACCENT if bold else C_TEXT)
        c.drawString(LABEL_X, ty, label)

        if units:
            c.setFont("Helvetica", 5.0 if small else 5.2)
            c.setFillColor(C_MUTED)
            c.drawCentredString(UNITS_X, ty, units)

        c.setFont("Helvetica-Bold" if bold else "Helvetica", amount_size)
        c.setFillColor(C_MUTED if small else C_ACCENT if bold else C_TEXT)
        c.drawRightString(AMT_X, ty, amount)

    body_bottom = BODY_TOP - len(data_rows) * ROW_H
    c.setStrokeColor(C_BORDER_LITE)
    c.setLineWidth(0.20)
    c.line(x + COL1_W, body_bottom, x + COL1_W, BODY_TOP)
    c.line(x + COL1_W + COL2_W, body_bottom, x + COL1_W + COL2_W, BODY_TOP)

    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.7)
    c.line(x, y, x + W, y)


# ─── Main generation function ─────────────────────────────────────────────────

def generate_pdf(excel_path: str, output_path: str,
                 company_name: str, pay_period: str,
                 progress_callback=None, log_callback=None) -> dict:
    """
    Generate the PDF.

    Args:
        excel_path:        Path to the input Excel file.
        output_path:       Path where the PDF will be saved.
        company_name:      Company name printed on each slip.
        pay_period:        Pay period string, e.g. "May-2026".
        progress_callback: Optional callable(percent: int) for progress updates.
        log_callback:      Optional callable(message: str) for log messages.

    Returns:
        dict with keys: success (bool), message (str), employee_count (int),
                        page_count (int), output_path (str)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    def progress(pct):
        if progress_callback:
            progress_callback(pct)

    try:
        log(f"📂  Reading Excel file: {os.path.basename(excel_path)}")
        progress(5)

        employees = read_employees(excel_path)
        if not employees:
            return {"success": False, "message": "No employee data found in Excel file.",
                    "employee_count": 0, "page_count": 0, "output_path": ""}

        log(f"✅  Loaded {len(employees)} employee record(s)")
        progress(15)

        PAGE_W, PAGE_H = A4
        usable_w = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT
        usable_h = PAGE_H - MARGIN_TOP  - MARGIN_BOTTOM
        cols          = int((usable_w + GAP) / (SLIP_WIDTH + GAP))
        rows_per_page = int((usable_h + GAP) / (SLIP_HEIGHT + GAP))
        per_page      = cols * rows_per_page
        total_pages   = (len(employees) - 1) // per_page + 1

        log(f"📐  Layout: {cols} columns × {rows_per_page} rows = {per_page} slips/page")
        log(f"📄  Total pages: {total_pages}")
        log(f"🖨️  Generating PDF …")
        progress(20)

        c = canvas.Canvas(output_path, pagesize=A4)
        c.setTitle(f"{company_name} Pay Slips – {pay_period}")
        c.setAuthor(f"{company_name} HR System")

        for idx, row in enumerate(employees):
            if idx % per_page == 0 and idx > 0:
                c.showPage()

            slot    = idx % per_page
            col_idx = slot % cols
            row_idx = slot // cols

            slip_x = MARGIN_LEFT + col_idx * (SLIP_WIDTH + GAP)
            slip_y = PAGE_H - MARGIN_TOP - (row_idx + 1) * SLIP_HEIGHT - row_idx * GAP

            draw_payslip(c, slip_x, slip_y, row, company_name, pay_period)

            pct = 20 + int(((idx + 1) / len(employees)) * 75)
            progress(pct)

        c.save()
        progress(100)
        log(f"✅  PDF saved successfully!")
        log(f"📁  Location: {output_path}")

        return {
            "success":        True,
            "message":        f"Successfully generated {len(employees)} pay slips across {total_pages} page(s).",
            "employee_count": len(employees),
            "page_count":     total_pages,
            "output_path":    output_path,
        }

    except FileNotFoundError as e:
        return {"success": False, "message": f"File not found: {e}",
                "employee_count": 0, "page_count": 0, "output_path": ""}
    except Exception as e:
        return {"success": False, "message": f"Error: {e}",
                "employee_count": 0, "page_count": 0, "output_path": ""}

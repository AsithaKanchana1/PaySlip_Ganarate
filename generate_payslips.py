#!/usr/bin/env python3
"""
Pay Slip Generator for New Lanka Clothing
=========================================
Reads employee data from Excel and generates a PDF with compact pay slips (5 cm × 12 cm)
fitted on A4 pages with 2 mm spacing between each slip.

Usage:
    python3 generate_payslips.py

Configuration:
    Edit the ── Configuration ── section below to change company name, pay period, etc.
"""

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import os
import sys

# ─── Configuration ─────────────────────────────────────────────────────────────
EXCEL_FILE  = "Excel/Slary_Slips.xlsx"
OUTPUT_PDF  = "PaySlips_Output.pdf"

COMPANY_NAME = "NEW LANKA CLOTHING"
PAY_PERIOD   = "May-2026"     # ← Change month/year here

# Pay slip dimensions
SLIP_WIDTH  = 5.0 * cm
SLIP_HEIGHT = 12.0 * cm
GAP         = 2 * mm

# A4 page margins
MARGIN_LEFT   = 0.7 * cm
MARGIN_RIGHT  = 0.7 * cm
MARGIN_TOP    = 0.7 * cm
MARGIN_BOTTOM = 0.7 * cm

# ─── Column indices (0-based) ──────────────────────────────────────────────────
# Match the exact order in your Excel sheet (row 1 = headers, row 2+ = data)
COL_EMP_NO       = 0
COL_NAME         = 1
COL_DEPARTMENT   = 2
COL_DESIGNATION  = 3
COL_BASIC        = 4
COL_ALLOWANCE    = 5   # Earnings: Allowance
COL_ATT_BONUS    = 6   # Earnings: Attendence Bonus
COL_FIXED_ALLOW  = 7
COL_INCENTIVE    = 8
COL_NORMAL_OT    = 9   # OT hours
COL_NORMAL_OT_AMT= 10  # OT amount
COL_DOUBLE_OT    = 11  # Double OT hours
COL_DOUBLE_OT_AMT= 12  # Double OT amount
COL_INCENTIVE2   = 13  # second incentive row
COL_GROSS        = 14
COL_SAL_ADV      = 15
COL_NO_PAY       = 16  # No-pay days
COL_NO_PAY_AMT   = 17  # No-pay amount
COL_ATT_BON_DED  = 18  # Deduction: Attendence Bonus
COL_ALLOW_DED    = 19  # Deduction: Allowance
COL_EPF8         = 20
COL_LATE         = 21  # Late count
COL_LATE_DED     = 22  # Late deduction
COL_WELFARE      = 23
COL_TOTAL_DED    = 24
COL_NET_SALARY   = 25
COL_EPF12        = 26
COL_ETF3         = 27

# ─── Colour palette ─────────────────────────────────────────────────────────────
C_HDR_BG      = colors.HexColor("#1a2744")  # deep navy header
C_HDR_TEXT    = colors.white
C_BORDER      = colors.HexColor("#2d3a6b")
C_BORDER_LITE = colors.HexColor("#c8d0e8")
C_ROW_ALT     = colors.HexColor("#f4f7fc")
C_GROSS_BG    = colors.HexColor("#fff8e1")  # amber
C_DED_BG      = colors.HexColor("#fde8e8")  # rose
C_NET_BG      = colors.HexColor("#e8f8ee")  # green
C_EPF_BG      = colors.HexColor("#f0f0f0")  # grey
C_TEXT        = colors.HexColor("#12192b")
C_MUTED       = colors.HexColor("#5a6080")
C_ACCENT      = colors.HexColor("#2d3a8c")


# ─── Helpers ────────────────────────────────────────────────────────────────────

def read_employees(path: str) -> list[list]:
    """Return list of rows (each row is a list of cell values), skipping header."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    employees = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        employees.append(list(row))
    wb.close()
    return employees


def g(row: list, col: int):
    """Safe column getter."""
    try:
        return row[col]
    except IndexError:
        return None


def fmt(val) -> str:
    """Format number to 2 decimal places with comma separator."""
    if val is None:
        return "0.00"
    try:
        return f"{float(val):,.2f}"
    except (ValueError, TypeError):
        return "0.00"


def fmt_units(val) -> str:
    """Format OT / Late / NoPay unit value; return empty string if zero."""
    if val is None or val == 0:
        return ""
    try:
        v = float(val)
        return f"{v:.2f}".rstrip("0").rstrip(".")
    except (ValueError, TypeError):
        return ""


# ─── Draw one pay slip ──────────────────────────────────────────────────────────

def draw_payslip(c: canvas.Canvas, x: float, y: float, row: list):
    """
    Draw a single pay slip.
    x, y  = bottom-left corner of the slip (ReportLab coordinate origin = bottom-left).
    """
    W = SLIP_WIDTH
    H = SLIP_HEIGHT

    # ── Outer border ────────────────────────────────────────────────────────────
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.7)
    c.rect(x, y, W, H, stroke=1, fill=0)

    # ── Header block ─────────────────────────────────────────────────────────────
    hdr_h = 1.62 * cm
    c.setFillColor(C_HDR_BG)
    c.rect(x, y + H - hdr_h, W, hdr_h, stroke=0, fill=1)

    # Company name
    c.setFillColor(C_HDR_TEXT)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString(x + W / 2, y + H - 0.38 * cm, COMPANY_NAME)

    # "Pay Sheet" label
    c.setFont("Helvetica-Bold", 6)
    c.drawCentredString(x + W / 2, y + H - 0.64 * cm, "Pay Sheet")

    # Period
    c.setFont("Helvetica-Bold", 6)
    c.drawCentredString(x + W / 2, y + H - 0.88 * cm, PAY_PERIOD)

    # Emp No
    c.setFont("Helvetica", 5.2)
    emp_no = g(row, COL_EMP_NO)
    emp_no_str = str(int(emp_no)) if emp_no else "-"
    c.drawString(x + 0.12 * cm, y + H - 1.13 * cm, f"Emp No   {emp_no_str}")

    # Name
    c.setFont("Helvetica-Bold", 5.5)
    name = g(row, COL_NAME) or ""
    c.drawCentredString(x + W / 2, y + H - 1.38 * cm, name)

    # Dept / Designation on one line
    c.setFont("Helvetica", 4.8)
    dept  = g(row, COL_DEPARTMENT)  or ""
    desig = g(row, COL_DESIGNATION) or ""
    c.drawString(x + 0.10 * cm, y + H - 1.58 * cm, f"Dept-{dept}   Desig-{desig}")

    # ── Build row data ────────────────────────────────────────────────────────────
    #  Each tuple: (label, unit_str, amount_str, style)
    #  style: "earn" | "gross" | "deduct" | "total_d" | "net" | "epf"
    data_rows = [
        ("Basic",             "",                              fmt(g(row, COL_BASIC)),        "earn"),
        ("Allowance",         "",                              fmt(g(row, COL_ALLOWANCE)),    "earn"),
        ("Attendance Bonus",  "",                              fmt(g(row, COL_ATT_BONUS)),    "earn"),
        ("Fixed Allowance",   "",                              fmt(g(row, COL_FIXED_ALLOW)),  "earn"),
        ("Incentive",         "",                              fmt(g(row, COL_INCENTIVE)),    "earn"),
        ("Normal OT",         fmt_units(g(row,COL_NORMAL_OT)),fmt(g(row, COL_NORMAL_OT_AMT)),"earn"),
        ("Double OT",         fmt_units(g(row,COL_DOUBLE_OT)),fmt(g(row, COL_DOUBLE_OT_AMT)),"earn"),
        ("Incentive",         "",                              fmt(g(row, COL_INCENTIVE2)),   "earn"),
        ("Gross Salary",      "",                              fmt(g(row, COL_GROSS)),        "gross"),
        ("Salary Advance",    "",                              fmt(g(row, COL_SAL_ADV)),      "deduct"),
        ("No Pay",            fmt_units(g(row,COL_NO_PAY)),   fmt(g(row, COL_NO_PAY_AMT)),   "deduct"),
        ("Attendance Bonus",  "",                              fmt(g(row, COL_ATT_BON_DED)),  "deduct"),
        ("Allowance",         "",                              fmt(g(row, COL_ALLOW_DED)),    "deduct"),
        ("E.P.F 8%",          "",                              fmt(g(row, COL_EPF8)),         "deduct"),
        ("Late",              fmt_units(g(row,COL_LATE)),     fmt(g(row, COL_LATE_DED)),     "deduct"),
        ("Welfare",           "",                              fmt(g(row, COL_WELFARE)),      "deduct"),
        ("Total Deduction",   "",                              fmt(g(row, COL_TOTAL_DED)),    "total_d"),
        ("Net Salary",        "",                              fmt(g(row, COL_NET_SALARY)),   "net"),
        ("E.P.F. 12%",        "",                              fmt(g(row, COL_EPF12)),        "epf"),
        ("E.T.F. 3%",         "",                              fmt(g(row, COL_ETF3)),         "epf"),
    ]

    # ── Body layout ───────────────────────────────────────────────────────────────
    ROW_H    = 0.302 * cm
    BODY_TOP = y + H - hdr_h        # top of body area
    COL1_W   = 2.30 * cm            # label column
    COL2_W   = 0.68 * cm            # units column
    COL3_W   = W - COL1_W - COL2_W  # amount column

    # Thin separator under header
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.35)
    c.line(x, BODY_TOP, x + W, BODY_TOP)

    for i, (label, units, amount, style) in enumerate(data_rows):
        ry = BODY_TOP - (i + 1) * ROW_H

        # Row background fill
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
            c.rect(x + 0.05, ry + 0.05, W - 0.1, ROW_H - 0.05, stroke=0, fill=1)

        # Row separator line
        c.setStrokeColor(C_BORDER_LITE)
        c.setLineWidth(0.18)
        c.line(x, ry, x + W, ry)

        # Text baseline
        ty = ry + 0.065 * cm

        # Label
        bold  = style in ("gross", "total_d", "net")
        small = style == "epf"
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 4.5 if small else 4.9)
        c.setFillColor(C_MUTED if small else C_ACCENT if bold else C_TEXT)
        c.drawString(x + 0.12 * cm, ty, label)

        # Units (center of col2)
        if units:
            c.setFont("Helvetica", 4.4)
            c.setFillColor(C_MUTED)
            c.drawCentredString(x + COL1_W + COL2_W / 2, ty, units)

        # Amount (right-aligned)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 4.5 if small else 4.9)
        c.setFillColor(C_MUTED if small else C_ACCENT if bold else C_TEXT)
        c.drawRightString(x + W - 0.10 * cm, ty, amount)

    # ── Vertical column separators in body ───────────────────────────────────────
    body_bottom = BODY_TOP - len(data_rows) * ROW_H
    c.setStrokeColor(C_BORDER_LITE)
    c.setLineWidth(0.20)
    c.line(x + COL1_W, body_bottom, x + COL1_W, BODY_TOP)
    c.line(x + COL1_W + COL2_W, body_bottom, x + COL1_W + COL2_W, BODY_TOP)

    # Bottom border of slip
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.7)
    c.line(x, y, x + W, y)


# ─── Layout & PDF generation ────────────────────────────────────────────────────

def generate_pdf(employees: list[list], output_path: str):
    PAGE_W, PAGE_H = A4

    usable_w = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT
    usable_h = PAGE_H - MARGIN_TOP  - MARGIN_BOTTOM

    cols          = int((usable_w + GAP) / (SLIP_WIDTH + GAP))
    rows_per_page = int((usable_h + GAP) / (SLIP_HEIGHT + GAP))
    per_page      = cols * rows_per_page

    print(f"  A4 page   : {PAGE_W/cm:.2f} × {PAGE_H/cm:.2f} cm")
    print(f"  Slip size : {SLIP_WIDTH/cm:.1f} × {SLIP_HEIGHT/cm:.1f} cm   gap: {GAP/mm:.0f} mm")
    print(f"  Layout    : {cols} cols × {rows_per_page} rows = {per_page} slips/page")
    print(f"  Employees : {len(employees)}")

    c = canvas.Canvas(output_path, pagesize=A4)
    c.setTitle(f"{COMPANY_NAME} Pay Slips – {PAY_PERIOD}")
    c.setAuthor("New Lanka Clothing HR System")

    for idx, row in enumerate(employees):
        if idx % per_page == 0 and idx > 0:
            c.showPage()

        slot    = idx % per_page
        col_idx = slot % cols
        row_idx = slot // cols

        slip_x = MARGIN_LEFT + col_idx * (SLIP_WIDTH + GAP)
        slip_y = PAGE_H - MARGIN_TOP - (row_idx + 1) * SLIP_HEIGHT - row_idx * GAP

        draw_payslip(c, slip_x, slip_y, row)

    c.save()
    total_pages = (len(employees) - 1) // per_page + 1 if employees else 1
    print(f"\n  ✅  Saved → {output_path}   ({total_pages} page(s))")


# ─── Entry point ────────────────────────────────────────────────────────────────

def main():
    base        = os.path.dirname(os.path.abspath(__file__))
    excel_path  = os.path.join(base, EXCEL_FILE)
    output_path = os.path.join(base, OUTPUT_PDF)

    if not os.path.exists(excel_path):
        print(f"❌  Excel file not found: {excel_path}")
        sys.exit(1)

    print("📂  Reading employee data …")
    employees = read_employees(excel_path)
    print(f"     {len(employees)} employee record(s) loaded")

    if not employees:
        print("⚠️   No employee data found.")
        sys.exit(0)

    print("🖨️  Generating PDF …")
    generate_pdf(employees, output_path)


if __name__ == "__main__":
    main()

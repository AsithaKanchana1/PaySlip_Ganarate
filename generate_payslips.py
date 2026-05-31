#!/usr/bin/env python3
"""
generate_payslips.py  —  New Lanka Clothing · Pay Slip Generator
================================================================
Standalone Python script — no GUI, no extra modules needed beyond
openpyxl and reportlab.

Developer: Asitha ❤️ Kanchana
GitHub   : https://github.com/AsithaKanchana1

USAGE
-----
1. Edit the ── Settings ── section below (company name, pay period, file paths)
2. Place your Excel salary file in the Excel/ folder
3. Run:
       python3 generate_payslips.py
4. Open/print PaySlips_Output.pdf

INSTALL DEPENDENCIES (once)
----------------------------
    pip install openpyxl reportlab

OUTPUT
------
    • Single PDF file with all employee pay slips
    • A4 page size
    • 3 columns × 2 rows = 6 pay slips per page
    • Each slip: 5 cm wide × 12 cm tall
    • 2 mm gap between slips
"""

import os
import sys
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors


# ══════════════════════════════════════════════════════════════════════════════
#  ── Settings ──  Edit these every month before running
# ══════════════════════════════════════════════════════════════════════════════

EXCEL_FILE   = "Excel/Slary_Slips.xlsx"   # path to your salary Excel file
OUTPUT_PDF   = "PaySlips_Output.pdf"       # output PDF filename
COMPANY_NAME = "NEW LANKA CLOTHING"        # printed on every slip
PAY_PERIOD   = "May-2026"                  # ← change this every month  e.g. "June-2026"


# ══════════════════════════════════════════════════════════════════════════════
#  ── Column indices ── (0-based, must match Excel column order)
# ══════════════════════════════════════════════════════════════════════════════
#  Column:  1=EmpNo  2=Name  3=Dept  4=Desig  5=Basic  6=Allow  7=AttBonus
#           8=FixedAllow  9=Incentive  10=NormOT  11=NormOTAmt  12=DblOT
#           13=DblOTAmt  14=Incentive2  15=Gross  16=SalAdv  17=NoPay
#           18=NoPayAmt  19=AttBonDed  20=AllowDed  21=EPF8  22=Late
#           23=LateDed  24=Welfare  25=TotalDed  26=NetSalary  27=EPF12  28=ETF3

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


# ══════════════════════════════════════════════════════════════════════════════
#  ── Layout constants ──
# ══════════════════════════════════════════════════════════════════════════════

SLIP_WIDTH    = 5.0 * cm
SLIP_HEIGHT   = 12.0 * cm
GAP           = 2 * mm

MARGIN_LEFT   = 0.7 * cm
MARGIN_RIGHT  = 0.7 * cm
MARGIN_TOP    = 0.7 * cm
MARGIN_BOTTOM = 0.7 * cm


# ══════════════════════════════════════════════════════════════════════════════
#  ── Colours ──
# ══════════════════════════════════════════════════════════════════════════════

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


# ══════════════════════════════════════════════════════════════════════════════
#  ── Helper functions ──
# ══════════════════════════════════════════════════════════════════════════════

def read_employees(excel_path: str) -> list:
    """Read all employee rows from the Excel file (skips header row 1)."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    employees = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:          # skip empty rows
            continue
        employees.append(list(row))
    wb.close()
    return employees


def _g(row: list, col: int):
    """Safe column value getter."""
    try:
        return row[col]
    except IndexError:
        return None


def _fmt(val) -> str:
    """Format a number as currency string with 2 decimal places."""
    if val is None:
        return "0.00"
    try:
        return f"{float(val):,.2f}"
    except (ValueError, TypeError):
        return "0.00"


def _fmt_units(val) -> str:
    """Format OT hours / late count — returns empty string if zero."""
    if val is None or val == 0:
        return ""
    try:
        v = float(val)
        s = f"{v:.2f}".rstrip("0").rstrip(".")
        return s
    except (ValueError, TypeError):
        return ""


# ══════════════════════════════════════════════════════════════════════════════
#  ── Draw a single pay slip ──
# ══════════════════════════════════════════════════════════════════════════════

def draw_payslip(c: canvas.Canvas, x: float, y: float, row: list):
    """
    Draw one pay slip at position (x, y) — bottom-left corner.
    ReportLab origin is bottom-left of the page.
    """
    W = SLIP_WIDTH
    H = SLIP_HEIGHT

    # ── Outer border ────────────────────────────────────────────────────────
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.7)
    c.rect(x, y, W, H, stroke=1, fill=0)

    # ── Header background ────────────────────────────────────────────────────
    hdr_h = 1.62 * cm
    c.setFillColor(C_HDR_BG)
    c.rect(x, y + H - hdr_h, W, hdr_h, stroke=0, fill=1)

    # Header text
    c.setFillColor(C_HDR_TEXT)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString(x + W / 2, y + H - 0.38 * cm, COMPANY_NAME)

    c.setFont("Helvetica-Bold", 6)
    c.drawCentredString(x + W / 2, y + H - 0.64 * cm, "Pay Sheet")
    c.drawCentredString(x + W / 2, y + H - 0.88 * cm, PAY_PERIOD)

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

    # ── Data rows ─────────────────────────────────────────────────────────────
    # (label, unit_value, amount_value, style)
    # style: "earn" | "gross" | "deduct" | "total_d" | "net" | "epf"
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

    ROW_H    = 0.302 * cm
    BODY_TOP = y + H - hdr_h
    COL1_W   = 2.30 * cm
    COL2_W   = 0.68 * cm

    # Separator under header
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.35)
    c.line(x, BODY_TOP, x + W, BODY_TOP)

    for i, (label, units, amount, style) in enumerate(data_rows):
        ry = BODY_TOP - (i + 1) * ROW_H

        # Row background
        if style == "gross":
            c.setFillColor(C_GROSS_BG)
        elif style == "total_d":
            c.setFillColor(C_DED_BG)
        elif style == "net":
            c.setFillColor(C_NET_BG)
        elif i % 2 == 0:
            c.setFillColor(C_ROW_ALT)

        if style in ("gross", "total_d", "net", "epf") or i % 2 == 0:
            c.setFillColor(C_ROW_ALT if (i % 2 == 0 and style == "earn") else
                           C_GROSS_BG if style == "gross" else
                           C_DED_BG if style == "total_d" else
                           C_NET_BG if style == "net" else
                           C_EPF_BG if style == "epf" else
                           C_ROW_ALT)
            c.rect(x + 0.05, ry + 0.05, W - 0.1, ROW_H - 0.05, stroke=0, fill=1)

        # Row separator line
        c.setStrokeColor(C_BORDER_LITE)
        c.setLineWidth(0.18)
        c.line(x, ry, x + W, ry)

        ty   = ry + 0.065 * cm
        bold  = style in ("gross", "total_d", "net")
        small = style == "epf"

        # Label
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 4.5 if small else 4.9)
        c.setFillColor(C_MUTED if small else C_ACCENT if bold else C_TEXT)
        c.drawString(x + 0.12 * cm, ty, label)

        # Units column
        if units:
            c.setFont("Helvetica", 4.4)
            c.setFillColor(C_MUTED)
            c.drawCentredString(x + COL1_W + COL2_W / 2, ty, units)

        # Amount (right-aligned)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 4.5 if small else 4.9)
        c.setFillColor(C_MUTED if small else C_ACCENT if bold else C_TEXT)
        c.drawRightString(x + W - 0.10 * cm, ty, amount)

    # Vertical column separators
    body_bottom = BODY_TOP - len(data_rows) * ROW_H
    c.setStrokeColor(C_BORDER_LITE)
    c.setLineWidth(0.20)
    c.line(x + COL1_W, body_bottom, x + COL1_W, BODY_TOP)
    c.line(x + COL1_W + COL2_W, body_bottom, x + COL1_W + COL2_W, BODY_TOP)

    # Bottom border
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(0.7)
    c.line(x, y, x + W, y)


# ══════════════════════════════════════════════════════════════════════════════
#  ── PDF layout & generation ──
# ══════════════════════════════════════════════════════════════════════════════

def generate_pdf(employees: list, output_path: str):
    """Lay out all pay slips on A4 pages and save to PDF."""
    PAGE_W, PAGE_H = A4
    usable_w = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT
    usable_h = PAGE_H - MARGIN_TOP  - MARGIN_BOTTOM

    cols          = int((usable_w + GAP) / (SLIP_WIDTH + GAP))
    rows_per_page = int((usable_h + GAP) / (SLIP_HEIGHT + GAP))
    per_page      = cols * rows_per_page
    total_pages   = (len(employees) - 1) // per_page + 1 if employees else 1

    print(f"  Page size  : A4  ({PAGE_W/cm:.1f} × {PAGE_H/cm:.1f} cm)")
    print(f"  Slip size  : {SLIP_WIDTH/cm:.0f} cm × {SLIP_HEIGHT/cm:.0f} cm   gap: {GAP/mm:.0f} mm")
    print(f"  Layout     : {cols} columns × {rows_per_page} rows = {per_page} slips/page")
    print(f"  Pages      : {total_pages}")

    c = canvas.Canvas(output_path, pagesize=A4)
    c.setTitle(f"{COMPANY_NAME} Pay Slips – {PAY_PERIOD}")
    c.setAuthor(f"{COMPANY_NAME} HR System")

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
    print(f"\n  ✅  Saved  →  {output_path}")


# ══════════════════════════════════════════════════════════════════════════════
#  ── Entry point ──
# ══════════════════════════════════════════════════════════════════════════════

def main():
    base_dir    = os.path.dirname(os.path.abspath(__file__))
    excel_path  = os.path.join(base_dir, EXCEL_FILE)
    output_path = os.path.join(base_dir, OUTPUT_PDF)

    print("━" * 52)
    print("  New Lanka Clothing — Pay Slip Generator (CLI)")
    print("━" * 52)
    print(f"  Company    : {COMPANY_NAME}")
    print(f"  Period     : {PAY_PERIOD}")
    print(f"  Excel file : {EXCEL_FILE}")
    print(f"  Output     : {OUTPUT_PDF}")
    print("━" * 52)

    # Validate Excel file exists
    if not os.path.exists(excel_path):
        print(f"\n  ❌  Excel file not found:")
        print(f"      {excel_path}")
        print(f"\n  → Place your Excel file at that path and try again.")
        print(f"  → See README.md for the required column format.")
        sys.exit(1)

    # Read data
    print("\n📂  Reading employee data …")
    employees = read_employees(excel_path)

    if not employees:
        print("  ⚠️   No employee data found in the Excel sheet.")
        print("       Make sure rows start from Row 2 and Emp No is filled.")
        sys.exit(0)

    print(f"  ✅  {len(employees)} employee record(s) loaded")

    # Generate PDF
    print("\n🖨️  Generating PDF …")
    generate_pdf(employees, output_path)

    print("\n" + "─" * 52)
    print(f"  Done!  {len(employees)} pay slips generated.")
    print(f"  Open {OUTPUT_PDF} and print on A4 at 100% scale.")
    print("─" * 52)
    print()
    print("  Asitha  ❤️  Kanchana")
    print("  🔗  github.com/AsithaKanchana1")
    print()


if __name__ == "__main__":
    main()

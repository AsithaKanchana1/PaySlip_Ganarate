#!/usr/bin/env python3
"""
generate_payslips.py  —  Command-line interface
================================================
Run from terminal when you don't need the GUI.
Edit the settings below, then run:
    python3 generate_payslips.py

For a graphical interface, run:
    python3 app.py
"""

import os
import sys
import getpass
from payslip_core import generate_pdf

# ─── Settings ─────────────────────────────────────────────────────────────────
EXCEL_FILE   = "Excel/Slary_Slips.xlsx"   # relative to this script's directory
OUTPUT_PDF   = "PaySlips_Output.pdf"       # relative to this script's directory
COMPANY_NAME = "NEW LANKA CLOTHING"
PAY_PERIOD   = "May-2026"                  # ← Change each month


# ─── Run ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    base        = os.path.dirname(os.path.abspath(__file__))
    excel_path  = os.path.join(base, EXCEL_FILE)
    output_path = os.path.join(base, OUTPUT_PDF)

    print("━" * 50)
    print("  New Lanka Clothing — Pay Slip Generator")
    print("━" * 50)

    if not os.path.exists(excel_path):
        print(f"❌  Excel file not found: {excel_path}")
        print(f"    Place your Excel file at: {excel_path}")
        sys.exit(1)

    pwd = getpass.getpass("Excel password (leave blank if none): ")
    ep = excel_path if not pwd else f"{excel_path}::{pwd}"

    result = generate_pdf(
        excel_path   = ep,
        output_path  = output_path,
        company_name = COMPANY_NAME,
        pay_period   = PAY_PERIOD,
        log_callback = print,
    )

    print("━" * 50)
    if result["success"]:
        print(f"✅  {result['message']}")
        print(f"📁  Output: {result['output_path']}")
    else:
        print(f"❌  {result['message']}")
        sys.exit(1)

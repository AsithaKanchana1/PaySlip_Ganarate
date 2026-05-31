# 🧾 New Lanka Clothing — Pay Slip Generator · Python CLI

> **Branch:** `python-cli` — Standalone Python script, no GUI, no installation needed.
> For the full desktop app with GUI and Windows installer, see the [`main`](https://github.com/AsithaKanchana1/PaySlip_Ganarate/tree/main) branch.

---

## What This Branch Contains

| File | Description |
|------|-------------|
| `generate_payslips.py` | **Single self-contained script** — edit settings at the top and run |
| `requirements.txt` | Only 2 dependencies: `openpyxl` + `reportlab` |
| `Excel/` | Place your salary Excel file here |
| `docs/` | Full documentation |

No GUI, no PyInstaller, no Windows installer — just Python.

---

## Quick Start

### 1 — Install dependencies (once)

```bash
pip install -r requirements.txt
```

### 2 — Add your Excel file

Copy your monthly salary file into the `Excel/` folder:
```
Excel/Slary_Slips.xlsx
```

### 3 — Set the pay period

Open `generate_payslips.py` in any text editor and change these lines near the top:

```python
# ── Settings ──
EXCEL_FILE   = "Excel/Slary_Slips.xlsx"   # path to your Excel file
OUTPUT_PDF   = "PaySlips_Output.pdf"       # output PDF name
COMPANY_NAME = "NEW LANKA CLOTHING"        # printed on every slip
PAY_PERIOD   = "May-2026"                  # ← change this every month
```

### 4 — Run the script

```bash
python3 generate_payslips.py
```

Output:
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  New Lanka Clothing — Pay Slip Generator (CLI)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Company    : NEW LANKA CLOTHING
  Period     : May-2026
  Excel file : Excel/Slary_Slips.xlsx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📂  Reading employee data …
  ✅  45 employee record(s) loaded

🖨️  Generating PDF …
  Page size  : A4  (21.0 × 29.7 cm)
  Slip size  : 5 cm × 12 cm   gap: 2 mm
  Layout     : 3 columns × 2 rows = 6 slips/page
  Pages      : 8
  ✅  Saved  →  PaySlips_Output.pdf
```

### 5 — Print the PDF

Open `PaySlips_Output.pdf` and print:
- Paper: **A4**
- Scale: **100% (Actual size)** — never "Fit to page"
- Orientation: **Portrait**

Each A4 page has **6 pay slips**. Cut and distribute.

---

## PDF Layout

```
┌──────────┬──────────┬──────────┐
│  Emp 001 │  Emp 002 │  Emp 003 │   ← 5 cm × 12 cm each
├──────────┼──────────┼──────────┤   ← 2 mm gap
│  Emp 004 │  Emp 005 │  Emp 006 │
└──────────┴──────────┴──────────┘
         6 slips per A4 page
```

---

## Excel Column Order

The script reads columns **by position**, not by name. Keep them in this exact order:

| # | Column | Example |
|---|--------|---------|
| 1 | Emp No | `411` |
| 2 | Name | `P.G.A.Kanchana` |
| 3 | Department | `Iron` |
| 4 | Designation | `Ironer` |
| 5 | Basic | `29000` |
| 6 | Allowance | `0` |
| 7 | Attendence Bonus | `10000` |
| 8 | Fixed Allowance | `0` |
| 9 | Incentive | `0` |
| 10 | Normal OT | `58.25` |
| 11 | Normal OT Amount | `8737.50` |
| 12 | Double OT | `8.50` |
| 13 | Double OT Amount | `1700` |
| 14 | Incentive | `0` |
| 15 | Gross Salary | `49437.50` |
| 16 | Salary Advance | `0` |
| 17 | No Pay | `1.5` |
| 18 | No Pay Amount | `0` |
| 19 | Attendence Bonus | `4750` |
| 20 | Allowance | `0` |
| 21 | E.P.F 8% | `2400` |
| 22 | Late | `5` |
| 23 | Late Deduction Amount | `400` |
| 24 | Walfare | `100` |
| 25 | Total Deduction | `7650` |
| 26 | Net Salary | `41787.50` |
| 27 | E.P.F. 12% | `3600` |
| 28 | E.T.F. 3% | `900` |

---

## Monthly Workflow

```
1. Update Excel file → save and close it
2. Edit PAY_PERIOD in generate_payslips.py
3. Run:  python3 generate_payslips.py
4. Print PaySlips_Output.pdf on A4 at 100%
5. Cut and distribute
```

---

## Data Privacy

The `.gitignore` keeps these files off GitHub:
- `Excel/*.xlsx` — private salary data
- `PaySlips_Output.pdf` — generated output

---

## Other Branches

| Branch | Description |
|--------|-------------|
| [`main`](https://github.com/AsithaKanchana1/PaySlip_Ganarate/tree/main) | Full desktop GUI app + Windows installer |
| `python-cli` | **This branch** — standalone Python script |

---

## Full Documentation

See the [`docs/`](./docs/) folder for detailed guides:

- [Getting Started](./docs/01_getting_started.md)
- [Excel File Format](./docs/03_excel_format.md)
- [PDF Output & Printing](./docs/05_pdf_output.md)
- [Troubleshooting](./docs/09_troubleshooting.md)

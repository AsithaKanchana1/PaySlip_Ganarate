# 🧾 New Lanka Clothing — Pay Slip Generator

> Automatically generates printer-ready PDF pay slips for all employees from a single Excel data file.  
> Each A4 page fits **6 pay slips** (3 columns × 2 rows) at **5 cm × 12 cm** each, separated by 2 mm gaps.

---

## 📁 Project Structure

```
Pay-Slip_Genarate/
│
├── Excel/                      ← 📂 Place your Excel data file HERE (not tracked by Git)
│   └── Slary_Slips.xlsx        ← Your monthly salary data (ignored by .gitignore)
│
├── generate_payslips.py        ← Main script — run this every month
├── index.html                  ← Column reference guide (open in browser)
├── PaySlips_Output.pdf         ← Generated output (ignored by .gitignore)
│
├── .gitignore                  ← Keeps Excel & PDF out of GitHub
└── README.md                   ← This file
```

---

## ⚙️ Requirements

Install the required Python libraries once:

```bash
pip install openpyxl reportlab
```

> If you are on **Arch Linux / externally managed Python**, use:
> ```bash
> pip install openpyxl reportlab --break-system-packages
> ```

---

## 📊 How to Prepare the Excel File

> ⚠️ **The Excel file contains private salary data — it is excluded from GitHub by `.gitignore`.**  
> You must place it manually on each machine before running the script.

### Step 1 — Create the `Excel` folder (first time only)

The folder already exists in the project. If it is missing, create it:

```bash
mkdir -p Excel
```

### Step 2 — Save your Excel file

Save (or copy) your monthly salary Excel sheet into the `Excel/` folder with **exactly this name**:

```
Excel/Slary_Slips.xlsx
```

> You can use a different filename — just update the `EXCEL_FILE` setting inside `generate_payslips.py`.

### Step 3 — Required Excel Column Order

Your Excel sheet **must have Row 1 as headers** and data starting from Row 2.  
The columns must appear in this exact order:

| # | Column Name | Description | Example |
|---|-------------|-------------|---------|
| 1 | `Emp No` | Employee number | `411` |
| 2 | `Name` | Full name | `P.G.A.Kanchana` |
| 3 | `Department` | Department | `Iron` |
| 4 | `Designation` | Job title | `Ironer` |
| 5 | `Basic` | Basic salary | `29000.00` |
| 6 | `Allowance` | Earnings allowance | `0.00` |
| 7 | `Attendence Bonus` | Attendance bonus (earnings) | `10000.00` |
| 8 | `Fixed Allowance` | Fixed allowance | `0.00` |
| 9 | `Incentive` | Incentive (earnings) | `0.00` |
| 10 | `Normal OT` | Normal OT hours | `58.25` |
| 11 | `Normal OT Amount` | Normal OT calculated amount | `8737.50` |
| 12 | `Double OT` | Double OT hours | `8.50` |
| 13 | `Double OT Amount` | Double OT amount | `1700.00` |
| 14 | `Incentive` | Second incentive row | `0.00` |
| 15 | `Gross Salary` | Total earnings | `49437.50` |
| 16 | `Salary Advance` | Advance deduction | `0.00` |
| 17 | `No Pay` | No-pay leave days | `1.50` |
| 18 | `No Pay Amount` | No-pay deduction amount | `0.00` |
| 19 | `Attendence Bonus` | Attendance bonus (deduction) | `4750.00` |
| 20 | `Allowance` | Allowance deduction | `0.00` |
| 21 | `E.P.F 8%` | EPF employee 8% | `2400.00` |
| 22 | `Late` | Late arrival count | `5.00` |
| 23 | `Late Deduction Amount` | Late deduction | `400.00` |
| 24 | `Walfare` | Welfare deduction | `100.00` |
| 25 | `Total Deduction` | Sum of all deductions | `7650.00` |
| 26 | `Net Salary` | Take-home pay | `41787.50` |
| 27 | `E.P.F. 12%` | EPF employer 12% | `3600.00` |
| 28 | `E.T.F. 3%` | ETF employer 3% | `900.00` |

> 💡 **Tip:** Use `Slary_Slips.xlsx` as your master template each month — just update the figures.

---

## 🖨️ Generate Pay Slips

### Step 1 — Set the pay period

Open `generate_payslips.py` in any text editor and update this line near the top:

```python
PAY_PERIOD = "May-2026"     # ← Change to current month and year
```

You can also change the company name if needed:

```python
COMPANY_NAME = "NEW LANKA CLOTHING"
```

### Step 2 — Run the script

```bash
python3 generate_payslips.py
```

You will see output like:

```
📂  Reading employee data …
     45 employee record(s) loaded
🖨️  Generating PDF …
  A4 page   : 21.00 × 29.70 cm
  Slip size : 5.0 × 12.0 cm   gap: 2 mm
  Layout    : 3 cols × 2 rows = 6 slips/page
  Employees : 45

  ✅  Saved → PaySlips_Output.pdf   (8 page(s))
```

### Step 3 — Print

Open `PaySlips_Output.pdf` and print with these settings:

| Setting | Value |
|---------|-------|
| Paper size | **A4** |
| Orientation | **Portrait** |
| Margins | **None / Minimal** |
| Scale | **100% (Actual size)** |
| Duplex | Optional |

After printing, cut each A4 sheet into 6 individual pay slips.

---

## 📐 Pay Slip Layout

```
┌─────────────────────────── A4 Page (21 × 29.7 cm) ───────────────────────────┐
│  ┌──────────┐  ┌──────────┐  ┌──────────┐                                    │
│  │ Emp 001  │  │ Emp 002  │  │ Emp 003  │   ← Row 1                          │
│  │ 5cm×12cm │  │ 5cm×12cm │  │ 5cm×12cm │                                    │
│  └──────────┘  └──────────┘  └──────────┘                                    │
│       ↕ 2mm gap between rows                                                  │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐                                    │
│  │ Emp 004  │  │ Emp 005  │  │ Emp 006  │   ← Row 2                          │
│  │ 5cm×12cm │  │ 5cm×12cm │  │ 5cm×12cm │                                    │
│  └──────────┘  └──────────┘  └──────────┘                                    │
└───────────────────────────────────────────────────────────────────────────────┘
                    6 pay slips per A4 page
```

---

## 🔒 Data Privacy & Git

The `.gitignore` file automatically **excludes** the following from GitHub:

| Item | Reason |
|------|--------|
| `Excel/*.xlsx` | Contains private employee salary data |
| `Excel/*.xls` | Same — older Excel format |
| `Excel/*.csv` | CSV exports |
| `PaySlips_Output.pdf` | Generated file with personal data |
| `*.pdf` | All PDF files |

> ✅ Only the **code** is pushed to GitHub — never the salary data.

### If you accidentally committed the Excel file

Run this to remove it from Git history:

```bash
git rm --cached Excel/Slary_Slips.xlsx
git commit -m "Remove Excel file from tracking"
git push
```

---

## 🔄 Monthly Workflow

```
1.  Update Excel → save to  Excel/Slary_Slips.xlsx
2.  Edit PAY_PERIOD in generate_payslips.py
3.  Run: python3 generate_payslips.py
4.  Print PaySlips_Output.pdf  (A4, 100%, no margins)
5.  Cut and distribute
```

---

## 🛠️ Troubleshooting

| Problem | Solution |
|---------|----------|
| `ModuleNotFoundError: openpyxl` | Run `pip install openpyxl reportlab` |
| `Excel file not found` | Check that `Excel/Slary_Slips.xlsx` exists |
| Pay slips cut off at bottom | Ensure print scale is 100%, no "fit to page" |
| Wrong month on slips | Update `PAY_PERIOD` in `generate_payslips.py` |
| Data shows `0.00` for all fields | Check Excel column order matches the table above |

---

## 📄 License

Internal use only — New Lanka Clothing HR Department.

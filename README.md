# 🧾 New Lanka Clothing — Pay Slip Generator

> Automatically generates printer-ready PDF pay slips for all employees from a single Excel data file.
> Available as a **click-and-use desktop app** for Windows — no technical knowledge needed.
> Linux and macOS are supported for developers.

---

## 🚀 Quick Start (For the Accountant — Windows)

> ✅ **No Python knowledge required.** Just install and use like any normal program.

### Option A — Use the Pre-Built Installer *(Easiest)*

1. Download **`PaySlipGenerator_Setup.exe`** (get it from your IT person or the Releases page)
2. Double-click to install → click Next → Finish
3. Open **Pay Slip Generator** from your Start Menu or Desktop
4. Copy your Excel salary file into the `Excel\` folder shown after installation
5. In the app: browse for the file, pick the month, click **Generate** ✅

---

### Option B — Run Directly with Python *(For developers)*

```bash
pip install -r requirements.txt
python app.py          # Opens the graphical desktop app
```

---

## 🖥️ App Screenshot Guide

```
┌─────────────────────────────────────────────────────┐
│  🧾  Pay Slip Generator     New Lanka Clothing       │
├─────────────────────────────────────────────────────┤
│  ⚙️ Company Settings                                │
│  Company Name: [ NEW LANKA CLOTHING              ]  │
│  Pay Period:   [ May ▼ ]  [ 2026 ]  → May-2026     │
├─────────────────────────────────────────────────────┤
│  📊 Excel Data File                                 │
│  [ Slary_Slips.xlsx              ] [ 📁 Browse ]   │
│  ✅ Slary_Slips.xlsx (62.0 KB)                      │
├─────────────────────────────────────────────────────┤
│  📁 Output PDF Location                             │
│  [ C:\Users\...\PaySlips_Output.pdf ] [ 📂 Choose ] │
├─────────────────────────────────────────────────────┤
│  [ 🖨️  Generate Pay Slips PDF                    ]  │
│  ████████████████████████░░░░  78%                  │
│  Generating pay slips …                             │
├─────────────────────────────────────────────────────┤
│  📋 Activity Log                                    │
│  📂 Reading Excel file: Slary_Slips.xlsx            │
│  ✅ Loaded 45 employee record(s)                    │
│  📐 Layout: 3 columns × 2 rows = 6 slips/page      │
│  🖨️ Generating PDF …                               │
│  ✅ PDF saved successfully!                         │
├─────────────────────────────────────────────────────┤
│  [ 📄  Open PDF                                  ]  │
└─────────────────────────────────────────────────────┘
```

---

## 👤 About & Developer Details

The desktop app includes an About panel with the following details:

- Developer: Asitha Kanchana
- GitHub: https://github.com/AsithaKanchana1
- Primary target: Windows 10/11
- Secondary support: Linux and macOS for development and testing

---

## 🏗️ Building the Windows Installer *(For IT / Developer)*

You need to do this once. After that, share the `PaySlipGenerator_Setup.exe` with the accountant.

### Prerequisites
- Windows 10 or 11
- [Python 3.11+](https://www.python.org/downloads/) — check "Add Python to PATH" during install
- [Inno Setup](https://jrsoftware.org/isdl.php) — free Windows installer creator

### Step-by-Step

**Step 1 — Clone the project**
```cmd
git clone https://github.com/AsithaKanchana1/PaySlip_Ganarate.git
cd PaySlip_Ganarate
```

**Step 2 — Build the .exe**

Double-click `build_windows.bat` or run in Command Prompt:
```cmd
build_windows.bat
```
This installs all dependencies and runs PyInstaller. Takes ~2 minutes.
Output: `dist\PaySlipGenerator\PaySlipGenerator.exe`

**Step 3 — Create the installer** *(optional but recommended)*

1. Open `installer.iss` in Inno Setup
2. Press **F9** (or Build → Compile)
3. Output: `Output\PaySlipGenerator_Setup.exe`

Share this single file with the accountant. ✅

---

## 📁 Project Structure

```
Pay-Slip_Genarate/
│
├── app.py                      ← 🖥️  Desktop GUI application (run this)
├── payslip_core.py             ← ⚙️  Core PDF generation logic
├── generate_payslips.py        ← 💻  Command-line interface (optional)
│
├── Excel/                      ← 📂  Place your Excel data file HERE
│   └── Slary_Slips.xlsx        ← (ignored by Git — add it manually)
│
├── requirements.txt            ← Python dependencies
├── build_windows.bat           ← One-click Windows .exe builder
├── installer.iss               ← Inno Setup — creates Setup.exe
│
├── .gitignore                  ← Keeps Excel & PDF out of GitHub
└── README.md                   ← This file
```

---

## 📊 Excel File Format

> ⚠️ **The Excel file is excluded from GitHub** (contains private salary data).
> Copy it manually to the `Excel\` folder on each machine.

### Required Columns (in this exact order, Row 1 = headers)

| # | Column Name | Example |
|---|-------------|---------|
| 1 | `Emp No` | `411` |
| 2 | `Name` | `P.G.A.Kanchana` |
| 3 | `Department` | `Iron` |
| 4 | `Designation` | `Ironer` |
| 5 | `Basic` | `29000.00` |
| 6 | `Allowance` | `0.00` |
| 7 | `Attendence Bonus` | `10000.00` |
| 8 | `Fixed Allowance` | `0.00` |
| 9 | `Incentive` | `0.00` |
| 10 | `Normal OT` | `58.25` |
| 11 | `Normal OT Amount` | `8737.50` |
| 12 | `Double OT` | `8.50` |
| 13 | `Double OT Amount` | `1700.00` |
| 14 | `Incentive` | `0.00` |
| 15 | `Gross Salary` | `49437.50` |
| 16 | `Salary Advance` | `0.00` |
| 17 | `No Pay` | `1.50` |
| 18 | `No Pay Amount` | `0.00` |
| 19 | `Attendence Bonus` | `4750.00` |
| 20 | `Allowance` | `0.00` |
| 21 | `E.P.F 8%` | `2400.00` |
| 22 | `Late` | `5.00` |
| 23 | `Late Deduction Amount` | `400.00` |
| 24 | `Walfare` | `100.00` |
| 25 | `Total Deduction` | `7650.00` |
| 26 | `Net Salary` | `41787.50` |
| 27 | `E.P.F. 12%` | `3600.00` |
| 28 | `E.T.F. 3%` | `900.00` |

---

## 📐 PDF Layout

- **Page size:** A4 (21 × 29.7 cm)
- **Each pay slip:** 5 cm wide × 12 cm tall
- **Gap between slips:** 2 mm
- **Slips per page:** 3 columns × 2 rows = **6 slips per A4 page**
- **Print settings:** A4, Portrait, 100% scale, No margins

---

## 🔄 Monthly Workflow (For the Accountant)

```
1. Open Excel → update figures → save file
2. Open "Pay Slip Generator" from Start Menu
3. Browse for the Excel file
4. Select the correct month and year
5. Click "Generate Pay Slips PDF"
6. Click "Open PDF" → Print on A4 paper
7. Cut each page into 6 slips and distribute ✅
```

---

## 🔒 Data Privacy

The `.gitignore` automatically keeps these off GitHub:

| Excluded | Reason |
|----------|--------|
| `Excel/*.xlsx` | Private salary data |
| `PaySlips_Output.pdf` | Contains personal employee data |
| `dist/`, `build/`, `Output/` | Build artifacts |

---

## 🛠️ Troubleshooting

| Problem | Solution |
|---------|----------|
| App won't start on Windows | Right-click → Run as Administrator |
| "No module named openpyxl" | Run `pip install -r requirements.txt` |
| Excel file not found | Check file is in the `Excel\` folder |
| PDF shows 0.00 for all values | Check Excel columns match the table above |
| Print cuts off slips | Set print scale to 100%, disable "Fit to page" |
| Build fails on step 3 | Make sure Python is in PATH; restart Command Prompt |

---

## 📄 License

Internal use only — New Lanka Clothing HR Department.

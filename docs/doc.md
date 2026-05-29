# 📚 New Lanka Clothing — Pay Slip Generator · Full Documentation

> **Version:** 1.0 &nbsp;|&nbsp; **Last Updated:** May 2026 &nbsp;|&nbsp; **Author:** New Lanka Clothing HR System

---

## 📖 Table of Contents

This folder contains the complete technical and user documentation for the **Pay Slip Generator** system used by New Lanka Clothing garment factory.

---

| # | Document | Audience | Description |
|---|----------|----------|-------------|
| 1 | [Getting Started](./01_getting_started.md) | Everyone | Installation & first-run guide |
| 2 | [Accountant User Guide](./02_accountant_guide.md) | Accountant / HR | How to use the app every month |
| 3 | [Excel File Format](./03_excel_format.md) | HR / Data Entry | Column layout & data rules |
| 4 | [Desktop App Reference](./04_app_gui.md) | All Users | Every button and field explained |
| 5 | [PDF Output & Printing](./05_pdf_output.md) | Accountant / HR | PDF layout, print settings, cutting |
| 6 | [Windows Installer Build](./06_windows_installer.md) | Developer / IT | How to build & distribute the .exe |
| 7 | [Developer Guide](./07_developer_guide.md) | Developer | Code architecture, extending the app |
| 8 | [Data Privacy & Git](./08_data_privacy.md) | Developer / IT | What stays private, .gitignore rules |
| 9 | [Troubleshooting](./09_troubleshooting.md) | Everyone | Common problems and fixes |

---

## 🏢 System Overview

```
┌──────────────────────────────────────────────────────────────────┐
│                  PAY SLIP GENERATOR — SYSTEM FLOW                │
│                                                                  │
│   ┌─────────────┐     ┌─────────────┐     ┌─────────────────┐  │
│   │ Excel File  │────▶│  app.py     │────▶│  PDF Output     │  │
│   │ (Salary     │     │  (GUI App)  │     │  (A4 pages,     │  │
│   │  Data)      │     │             │     │   6 slips/page) │  │
│   └─────────────┘     └──────┬──────┘     └─────────────────┘  │
│                              │                                   │
│                       ┌──────▼──────┐                           │
│                       │payslip_core │                           │
│                       │.py (Engine) │                           │
│                       └─────────────┘                           │
└──────────────────────────────────────────────────────────────────┘
```

### Key Facts

| Property | Value |
|----------|-------|
| Pay slip size | 5 cm × 12 cm |
| Slips per A4 page | 6 (3 columns × 2 rows) |
| Gap between slips | 2 mm |
| Output format | Single PDF file |
| Input format | Microsoft Excel (.xlsx) |
| Supported OS | Windows 10/11 (installer), Linux, macOS |

---

## 📁 Project File Map

```
Pay-Slip_Genarate/
│
├── docs/                        ← 📚 You are here
│   ├── doc.md                   ← Master documentation index
│   ├── 01_getting_started.md
│   ├── 02_accountant_guide.md
│   ├── 03_excel_format.md
│   ├── 04_app_gui.md
│   ├── 05_pdf_output.md
│   ├── 06_windows_installer.md
│   ├── 07_developer_guide.md
│   ├── 08_data_privacy.md
│   └── 09_troubleshooting.md
│
├── Excel/                       ← Salary data (NOT on GitHub)
│   └── Slary_Slips.xlsx
│
├── app.py                       ← Desktop GUI application
├── payslip_core.py              ← Core PDF generation engine
├── generate_payslips.py         ← Command-line interface
├── requirements.txt             ← Python dependencies
├── build_windows.bat            ← Windows .exe builder
├── installer.iss                ← Inno Setup installer script
├── .gitignore                   ← Git privacy rules
└── README.md                    ← Project overview
```

---

## ⚡ Quick Reference

### For the Accountant
```
Every month:
  1. Open "Pay Slip Generator" from Start Menu
  2. Browse → select Excel salary file
  3. Choose month and year
  4. Click "Generate Pay Slips PDF"
  5. Click "Open PDF" → Print → Cut → Distribute
```

### For the Developer
```bash
# Run GUI
python app.py

# Run CLI
python generate_payslips.py

# Install dependencies
pip install -r requirements.txt

# Build Windows exe (on Windows)
build_windows.bat
```

---

*For detailed instructions on any topic, click the relevant document in the table above.*

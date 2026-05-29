# 04 · Desktop App Reference (GUI Guide)

> **Audience:** All users
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

The **Pay Slip Generator** desktop app (`app.py`) provides a graphical interface for generating pay slips. This document explains every section, field, and button in the app.

---

## Launching the App

| Method | Command / Action |
|--------|-----------------|
| Windows (installed) | Start Menu → **Pay Slip Generator** |
| Windows (portable) | Double-click `PaySlipGenerator.exe` |
| Python (any OS) | `python3 app.py` in the project folder |

---

## App Window — Full Reference

### Header Bar

```
┌──────────────────────────────────────────────────────┐
│  🧾  Pay Slip Generator          New Lanka Clothing  │
└──────────────────────────────────────────────────────┘
```

| Element | Description |
|---------|-------------|
| 🧾 Title | App name |
| Right label | Company branding |

---

### Section 1 — Company Settings

```
⚙️ Company Settings
───────────────────────────────────────────
Company Name
[ NEW LANKA CLOTHING                      ]

Pay Period
[ May ▼ ]  [ 2026 ]   →  May-2026
```

#### Company Name Field
- **What it does:** The text entered here is printed as the heading on every pay slip
- **Default value:** `NEW LANKA CLOTHING`
- **How to change:** Click inside the field and edit the text
- **Remembered:** Yes — the app saves this and pre-fills it next time

#### Month Dropdown
- **What it does:** Sets the month printed on each pay slip
- **Options:** January through December
- **How to change:** Click the blue dropdown → select a month
- **Remembered:** Yes

#### Year Field
- **What it does:** Sets the year printed on each pay slip
- **Format:** 4-digit year (e.g. `2026`)
- **How to change:** Click and type a new year
- **Remembered:** Yes

#### Pay Period Preview
- Shows the formatted pay period that will appear on slips
- Updates instantly as you change month or year
- Format: `Month-Year` (e.g. `May-2026`)

---

### Section 2 — Excel Data File

```
📊 Excel Data File
───────────────────────────────────────────
Select the monthly salary Excel file (*.xlsx)

[ Excel/Slary_Slips.xlsx              ] [ 📁 Browse ]
✅  Slary_Slips.xlsx  (62.0 KB)
```

#### File Path Display
- Shows the full path to the currently selected Excel file
- Read-only — you cannot type in this field directly
- Use the Browse button to change the file

#### Browse Button (📁 Browse)
- Opens a file picker dialog
- Filters to show only `.xlsx` and `.xls` files
- After selection, the file path updates and a confirmation message appears
- **Also automatically sets** the output PDF path to the same folder as the Excel file

#### Status Message
- ✅ Green — file selected successfully (shows filename and size)
- ❌ Red — file not found or cannot be read

---

### Section 3 — Output PDF Location

```
📁 Output PDF Location
───────────────────────────────────────────
Where to save the generated PDF file

[ C:\Users\Accountant\Desktop\PaySlips_Output.pdf ] [ 📂 Choose ]
```

#### Output Path Field
- Shows where the PDF will be saved
- **Default:** Same folder as the selected Excel file, named `PaySlips_Output.pdf`
- You can edit this path directly by clicking on it
- **Remembered:** Yes

#### Choose Button (📂 Choose)
- Opens a Save As dialog
- Lets you choose a different folder and filename
- Default suggested name: `PaySlips_Output.pdf`

> ⚠️ If a PDF already exists at the output path, it will be **overwritten** without warning. Make sure to keep a copy if you need the previous month's PDF.

---

### Generate Button

```
[ 🖨️  Generate Pay Slips PDF ]
```

- **Colour:** Green
- **What it does:** Reads the Excel file and generates the PDF
- **Disabled while:** Generation is in progress
- **Validation:** Checks that an Excel file is selected, output path is set, company name is filled, and year is valid — shows an error message if anything is missing

---

### Progress Bar

```
████████████████████░░░░░░░░░░  68%
Generating pay slips …
```

- Fills from 0% to 100% as each employee's slip is generated
- The status text below it shows the current step

| Progress % | Step |
|-----------|------|
| 0–5% | Starting |
| 5–15% | Reading Excel file |
| 15–20% | Calculating layout |
| 20–95% | Drawing each slip |
| 95–100% | Saving PDF |

---

### Activity Log

```
📋 Activity Log
───────────────────────────────────────────
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Company   : NEW LANKA CLOTHING
  Period    : May-2026
  Excel     : Slary_Slips.xlsx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📂 Reading Excel file: Slary_Slips.xlsx
✅  Loaded 45 employee record(s)
📐  Layout: 3 columns × 2 rows = 6 slips/page
📄  Total pages: 8
🖨️  Generating PDF …
✅  PDF saved successfully!
📁  Location: C:\Users\...\PaySlips_Output.pdf
```

- Shows a real-time log of everything the app is doing
- Scrollable — scroll up to see earlier messages
- Cleared automatically at the start of each generation
- Useful for diagnosing problems

---

### Open PDF Button

```
[ 📄  Open PDF ]
```

- **Colour:** Blue
- **State:** Disabled until a PDF is successfully generated
- **What it does:** Opens the generated PDF in your default PDF viewer (Adobe Reader, Edge, Chrome, etc.)
- If the file cannot be found, shows a warning message

---

## Settings Persistence

The app automatically saves and restores these settings between sessions:

| Setting | Saved? |
|---------|--------|
| Company name | ✅ Yes |
| Last selected month | ✅ Yes |
| Last used year | ✅ Yes |
| Last Excel file path | ✅ Yes (if file still exists) |
| Last output PDF path | ✅ Yes |

Settings are saved to: `~/.payslip_settings.json`
(On Windows: `C:\Users\YourName\.payslip_settings.json`)

---

## Success Popup

After successful generation, a popup appears:

```
┌─────────────────────────────────────────┐
│  ✅  Success                            │
│                                         │
│  Pay slips generated successfully!      │
│                                         │
│  👥  Employees : 45                     │
│  📄  Pages     : 8                      │
│                                         │
│  📁  Saved to:                          │
│  C:\Users\...\PaySlips_Output.pdf       │
│                                         │
│               [ OK ]                    │
└─────────────────────────────────────────┘
```

Click **OK** to dismiss. Then use the **📄 Open PDF** button to view and print.

---

## Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| "Please browse and select the Excel salary file first." | No file selected | Click Browse and pick the Excel file |
| "File Not Found" | Excel file was moved or deleted | Browse again and locate the file |
| "Please choose where to save the PDF." | Output path is empty | Click Choose and set a save location |
| "Please enter the company name." | Company name field is empty | Type the company name |
| "Invalid Year" | Year is not 4 digits | Enter a year like `2026` |
| "Generation Failed" | Error during PDF creation | See the Activity Log for details |

---

## Related Documents

- [Accountant User Guide](./02_accountant_guide.md) — step-by-step monthly workflow
- [Excel File Format](./03_excel_format.md) — data requirements
- [PDF Output & Printing](./05_pdf_output.md) — print settings
- [Troubleshooting](./09_troubleshooting.md) — fixing problems

# 02 · Accountant User Guide

> **Audience:** Accountant / HR staff — monthly pay slip generation
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

This guide explains how to generate pay slips every month using the **Pay Slip Generator** desktop app. No technical knowledge is needed — just follow these steps each pay period.

---

## Monthly Workflow — Step by Step

### Step 1 — Update the Excel Salary File

Before opening the app, make sure your Excel salary file is up to date for the current month.

1. Open `Slary_Slips.xlsx` in Microsoft Excel
2. Update all salary figures, OT hours, deductions, etc. for each employee
3. **Save and close** the Excel file
4. Make sure the file is saved in the correct location (see [Excel File Format](./03_excel_format.md))

> ⚠️ **Important:** The app cannot read the Excel file if it is still open in Excel at the same time. Always close Excel before generating pay slips.

---

### Step 2 — Open the Pay Slip Generator App

- Click **Start Menu** → search **Pay Slip Generator** → click to open
- OR double-click the desktop shortcut (if you created one during installation)

The app window will open:

```
┌─────────────────────────────────────────────────────┐
│  🧾  Pay Slip Generator     New Lanka Clothing       │
└─────────────────────────────────────────────────────┘
```

---

### Step 3 — Check the Company Name

The **Company Name** field should already show `NEW LANKA CLOTHING`.

- If it is correct → leave it as is
- If it needs changing → click the field and type the correct name

> 💡 The app remembers your last-used company name automatically.

---

### Step 4 — Set the Pay Period

Select the correct **month** and **year** for this pay period.

| Field | How to change |
|-------|---------------|
| Month | Click the dropdown → select the month |
| Year | Click the year box → type the year (e.g. `2026`) |

The preview on the right shows what will be printed on the slips:
```
→  May-2026
```

> ⚠️ **Always double-check** the month and year before generating. This is printed on every pay slip and cannot be changed after the PDF is created without regenerating.

---

### Step 5 — Select the Excel File

1. Click the **📁 Browse** button
2. Navigate to where your `Slary_Slips.xlsx` file is saved
3. Click on the file → click **Open**

The app will confirm the file is loaded:
```
✅  Slary_Slips.xlsx  (62.0 KB)
```

> 💡 The app remembers the last Excel file you used. Next month, it will pre-fill the same location — you just need to verify it is still correct.

---

### Step 6 — Check the Output Location

The **Output PDF Location** shows where the generated PDF will be saved.

- By default it saves next to your Excel file as `PaySlips_Output.pdf`
- To change it: click **📂 Choose** → browse to a different folder → give the file a name → click **Save**

> 💡 **Recommended:** Keep the default location (same folder as your Excel file) so you can always find the PDF easily.

---

### Step 7 — Generate the Pay Slips

Click the green button:
```
[ 🖨️  Generate Pay Slips PDF ]
```

You will see:
- A **progress bar** filling up (0% → 100%)
- A **live log** showing what the app is doing
- A **success popup** when done

Example log:
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Company   : NEW LANKA CLOTHING
  Period    : May-2026
  Excel     : Slary_Slips.xlsx
  Output    : C:\Users\...\PaySlips_Output.pdf
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📂 Reading Excel file: Slary_Slips.xlsx
✅ Loaded 45 employee record(s)
📐 Layout: 3 columns × 2 rows = 6 slips/page
📄 Total pages: 8
🖨️ Generating PDF …
✅ PDF saved successfully!
```

---

### Step 8 — Open and Print the PDF

1. Click the blue **📄 Open PDF** button — the PDF opens automatically
2. In your PDF viewer (Adobe Reader, Edge, Chrome), press **Ctrl+P** to print
3. Set print settings:

| Setting | Value |
|---------|-------|
| Printer | Your office printer |
| Paper size | **A4** |
| Orientation | **Portrait** |
| Scale / Size | **100% (Actual Size)** — do NOT use "Fit to page" |
| Margins | **None** or **Minimum** |
| Copies | As needed |

4. Click **Print**

> ⚠️ **Always print at 100% scale.** If you use "Fit to page" or any other scaling, the pay slip sizes will be wrong and they will not cut correctly.

---

### Step 9 — Cut and Distribute

Each A4 page contains **6 pay slips** arranged in 3 columns and 2 rows.

```
┌──────────┬──────────┬──────────┐
│  Emp 001 │  Emp 002 │  Emp 003 │  ← Cut here (horizontal)
├──────────┼──────────┼──────────┤
│  Emp 004 │  Emp 005 │  Emp 006 │
└──────────┴──────────┴──────────┘
     ↑           ↑          ↑
  Cut here    Cut here   Cut here (vertical)
```

- Cut **horizontally** across the middle of each page (between row 1 and row 2)
- Cut **vertically** between each column
- Each slip is **5 cm wide** and **12 cm tall**
- There is a **2 mm gap** between slips to guide your cutting

---

## Quick Reference Card

Print this and keep it at your desk:

```
╔══════════════════════════════════════╗
║  PAY SLIP GENERATOR — MONTHLY STEPS ║
╠══════════════════════════════════════╣
║  1. Update Excel file & close it    ║
║  2. Open Pay Slip Generator app     ║
║  3. Check company name              ║
║  4. Set month and year              ║
║  5. Browse → select Excel file      ║
║  6. Click Generate                  ║
║  7. Print at 100% scale on A4       ║
║  8. Cut and distribute              ║
╚══════════════════════════════════════╝
```

---

## Frequently Asked Questions

**Q: The PDF was generated but some employees have 0.00 for all amounts.**
A: The Excel file column order may be wrong. See [Excel File Format](./03_excel_format.md).

**Q: I selected the wrong month. Can I fix it?**
A: Yes — change the month in the app and click Generate again. The old PDF will be replaced.

**Q: The app shows an error when I click Generate.**
A: Check that the Excel file is closed in Excel and that you selected the correct file. See [Troubleshooting](./09_troubleshooting.md).

**Q: How do I add or remove employees?**
A: Add or remove rows in the Excel file. The app automatically picks up all rows with data.

**Q: Can I generate pay slips for just one employee?**
A: Not directly from the app — but you can temporarily save a separate Excel file with only that employee's row.

---

## Related Documents

- [Excel File Format](./03_excel_format.md) — column layout and data entry rules
- [PDF Output & Printing](./05_pdf_output.md) — detailed print settings
- [App Reference](./04_app_gui.md) — every button explained
- [Troubleshooting](./09_troubleshooting.md) — what to do when something goes wrong

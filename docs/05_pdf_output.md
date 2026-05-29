# 05 · PDF Output & Printing

> **Audience:** Accountant / HR
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

The Pay Slip Generator produces a single PDF file containing all employees' pay slips, laid out to be printed on A4 paper and then cut into individual slips.

---

## Pay Slip Dimensions

| Property | Value |
|----------|-------|
| Width | 5 cm |
| Height | 12 cm |
| Gap between slips | 2 mm |
| Page size | A4 (21 × 29.7 cm) |
| Slips per page | 6 (3 columns × 2 rows) |
| Orientation | Portrait |

---

## Page Layout Diagram

```
◄─────────────── A4 Page (21 cm) ───────────────►
▲  ┌──────────┐  ┌──────────┐  ┌──────────┐
│  │          │  │          │  │          │
│  │  Emp 001 │  │  Emp 002 │  │  Emp 003 │
│  │  5×12 cm │  │  5×12 cm │  │  5×12 cm │
│  │          │  │          │  │          │
A  └──────────┘  └──────────┘  └──────────┘
4     ← 2mm →                    (2mm gap)
▼  ┌──────────┐  ┌──────────┐  ┌──────────┐
│  │          │  │          │  │          │
│  │  Emp 004 │  │  Emp 005 │  │  Emp 006 │
│  │  5×12 cm │  │  5×12 cm │  │  5×12 cm │
│  │          │  │          │  │          │
▼  └──────────┘  └──────────┘  └──────────┘
```

---

## What is Printed on Each Slip

Every pay slip contains the following information in this layout:

```
┌────────────────────────────────────┐
│       NEW LANKA CLOTHING           │  ← Company name (header)
│           Pay Sheet                │
│           May-2026                 │  ← Pay period
│ Emp No   411                       │
│       P.G.A.Kanchana               │  ← Employee name
│ Dept-Iron   Desig-Ironer           │
├──────────────────┬────────┬────────┤
│ Basic            │        │ 29,000 │
│ Allowance        │        │   0.00 │
│ Attendance Bonus │        │ 10,000 │
│ Fixed Allowance  │        │   0.00 │
│ Incentive        │        │   0.00 │
│ Normal OT        │  58.25 │  8,737 │  ← Hours shown in middle
│ Double OT        │   8.50 │  1,700 │
│ Incentive        │        │   0.00 │
├──────────────────┴────────┴────────┤
│ Gross Salary               49,437  │  ← Highlighted amber
├──────────────────┬────────┬────────┤
│ Salary Advance   │        │   0.00 │
│ No Pay           │   1.50 │   0.00 │  ← Days shown in middle
│ Attendance Bonus │        │  4,750 │
│ Allowance        │        │   0.00 │
│ E.P.F 8%         │        │  2,400 │
│ Late             │   5.00 │    400 │  ← Count shown in middle
│ Welfare          │        │    100 │
├──────────────────┴────────┴────────┤
│ Total Deduction             7,650  │  ← Highlighted pink
├────────────────────────────────────┤
│ Net Salary                 41,787  │  ← Highlighted green (bold)
├────────────────────────────────────┤
│ E.P.F. 12%                  3,600  │  ← Grey (employer contrib)
│ E.T.F. 3%                     900  │
└────────────────────────────────────┘
```

---

## Output File

| Property | Value |
|----------|-------|
| Default filename | `PaySlips_Output.pdf` |
| Default location | Same folder as the Excel file |
| Format | PDF 1.4 compatible |
| Fonts | Helvetica (embedded) |
| Colour | Full colour (colour printer recommended) |

---

## Print Settings

### In Adobe Acrobat Reader

1. Open the PDF → Press `Ctrl+P`
2. Set these options:

| Setting | Value |
|---------|-------|
| Printer | Your office printer |
| Page sizing | **Actual size** (NOT Fit, Shrink, or Scale) |
| Paper | A4 |
| Orientation | Portrait (auto) |
| Pages | All |

3. Click **Print**

### In Microsoft Edge (PDF viewer)

1. Open PDF → Click Print icon (or `Ctrl+P`)
2. Set:
 - **Layout:** Portrait
 - **Paper size:** A4
 - **Scale:** Custom → enter **100**
3. Click **Print**

### In Google Chrome

1. Open PDF → `Ctrl+P`
2. Set:
 - **Paper size:** A4
 - **Scale:** **100** (default)
 - Uncheck **Fit to page** if shown
3. Click **Print**

> ⚠️ The most important setting across all viewers is **do not scale / fit to page**. Always print at exactly 100% so the 5 cm × 12 cm dimensions are accurate.

---

## Cutting the Slips

### Tools needed
- Paper guillotine / cutting machine (recommended for large batches)
- OR scissors and a ruler

### Cutting guide

**For each A4 page:**

1. **First cut — horizontal** (across the width of the page)
   - Cut between row 1 and row 2
   - The 2 mm gap between rows is your guide line
   - This gives you two strips, each 12 cm tall

2. **Second cut — vertical** (down the height of each strip)
   - Cut between column 1 and column 2, then between column 2 and column 3
   - The 2 mm gap between columns is your guide line
   - This gives you 3 individual slips per strip

**Result:** 6 individual pay slips (5 cm × 12 cm each) per A4 page.

---

## How Many Pages Will be Generated?

| Employees | Pages Needed |
|-----------|--------------|
| 1–6 | 1 page |
| 7–12 | 2 pages |
| 13–18 | 3 pages |
| 19–24 | 4 pages |
| 25–30 | 5 pages |
| Every 6 more | +1 page |

Formula: `Pages = CEILING(Employee Count / 6)`

---

## Colour Coding on Each Slip

| Colour | Row |
|--------|-----|
| 🟨 Amber/Yellow | Gross Salary row |
| 🟥 Light Red | Total Deduction row |
| 🟩 Light Green | Net Salary row |
| ⬜ Light Grey | EPF 12% and ETF 3% rows |
| 🔵 Alternating light blue | Regular earnings/deduction rows |

---

## Related Documents

- [Accountant User Guide](./02_accountant_guide.md) — full monthly workflow
- [App Reference](./04_app_gui.md) — output path settings
- [Troubleshooting](./09_troubleshooting.md) — print problems

# 03 · Excel File Format

> **Audience:** HR staff / Data Entry / Developer
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

The Pay Slip Generator reads salary data from a Microsoft Excel file (`.xlsx` format). This document explains the exact structure the file must follow.

---

## File Location

Save the Excel file here inside the project folder:

```
Pay-Slip_Genarate/
└── Excel/
    └── Slary_Slips.xlsx     ← exact filename expected by default
```

> 💡 You can use a different filename or folder — just browse to it in the app when generating.

> ⚠️ **Never rename or move the file while it is selected in the app.** Always close the app first.

---

## File Structure Rules

| Rule | Requirement |
|------|-------------|
| Format | `.xlsx` (Excel 2007 or newer) |
| Active sheet | The **first sheet** (Sheet1) is read automatically |
| Row 1 | **Header row** — column names (must be present) |
| Row 2 onwards | One row per employee |
| Empty rows | Automatically skipped |
| Column order | Must match the exact order below |

---

## Column Reference — All 28 Columns

> ⚠️ **Column position matters.** The app reads columns by their position (1st, 2nd, 3rd…), not by their header name. Keep them in this exact order.

### Section 1 — Employee Information (Columns 1–4)

| # | Header | Data Type | Notes |
|---|--------|-----------|-------|
| 1 | `Emp No` | Number | Unique employee number (e.g. `411`) |
| 2 | `Name` | Text | Full name as it should appear on the slip |
| 3 | `Department` | Text | Department name (e.g. `Iron`, `Packing`) |
| 4 | `Designation` | Text | Job title (e.g. `Ironer`, `Packer`) |

### Section 2 — Earnings (Columns 5–14)

| # | Header | Data Type | Notes |
|---|--------|-----------|-------|
| 5 | `Basic` | Number | Basic monthly salary |
| 6 | `Allowance` | Number | General earnings allowance |
| 7 | `Attendence Bonus` | Number | Attendance bonus (earnings side) |
| 8 | `Fixed Allowance` | Number | Fixed monthly allowance |
| 9 | `Incentive` | Number | First incentive figure |
| 10 | `Normal OT` | Number | Normal overtime **hours** (not amount) |
| 11 | `Normal OT Amount` | Number | Normal OT calculated **amount** |
| 12 | `Double OT` | Number | Double overtime **hours** |
| 13 | `Double OT Amount` | Number | Double OT calculated **amount** |
| 14 | `Incentive` | Number | Second incentive figure |

> 💡 Note that `Incentive` appears twice (columns 9 and 14) and `Attendence Bonus` appears twice in the sheet (columns 7 and 19 — one for earnings, one for deductions). This is by design to match the original pay slip format.

### Section 3 — Gross (Column 15)

| # | Header | Data Type | Notes |
|---|--------|-----------|-------|
| 15 | `Gross Salary` | Number | Sum of all earnings. Should equal: Basic + Allowances + Bonuses + OT amounts |

### Section 4 — Deductions (Columns 16–24)

| # | Header | Data Type | Notes |
|---|--------|-----------|-------|
| 16 | `Salary Advance` | Number | Any advance already paid |
| 17 | `No Pay` | Number | No-pay leave **days** (not amount) |
| 18 | `No Pay Amount` | Number | No-pay deduction **amount** |
| 19 | `Attendence Bonus` | Number | Attendance bonus deduction (second appearance) |
| 20 | `Allowance` | Number | Allowance deduction (second appearance) |
| 21 | `E.P.F 8%` | Number | Employee EPF contribution (8% of basic) |
| 22 | `Late` | Number | Number of late arrivals |
| 23 | `Late Deduction Amount` | Number | Late deduction **amount** |
| 24 | `Walfare` | Number | Welfare fund deduction |

> ⚠️ Note the spelling: `Walfare` (not `Welfare`) — this matches the original Excel sheet header.

### Section 5 — Totals & Employer Contributions (Columns 25–28)

| # | Header | Data Type | Notes |
|---|--------|-----------|-------|
| 25 | `Total Deduction` | Number | Sum of all deductions (columns 16–24) |
| 26 | `Net Salary` | Number | Gross Salary − Total Deduction |
| 27 | `E.P.F. 12%` | Number | Employer EPF contribution (12% of basic) |
| 28 | `E.T.F. 3%` | Number | Employer ETF contribution (3% of basic) |

---

## Sample Data Row

| Emp No | Name | Dept | Desig | Basic | Allow | AttBonus | FixedAllow | Incentive | NormOT | NormOT Amt | DblOT | DblOT Amt | Incentive2 | Gross |
|--------|------|------|-------|-------|-------|----------|------------|-----------|--------|------------|-------|-----------|------------|-------|
| 411 | P.G.A.Kanchana | Iron | Ironer | 29000 | 0 | 10000 | 0 | 0 | 58.25 | 8737.50 | 8.50 | 1700 | 0 | 49437.50 |

| SalAdv | NoPay | NoPayAmt | AttBonDed | AllowDed | EPF8 | Late | LateDed | Welfare | TotalDed | NetSal | EPF12 | ETF3 |
|--------|-------|----------|-----------|----------|------|------|---------|---------|----------|--------|-------|------|
| 0 | 1.5 | 0 | 4750 | 0 | 2400 | 5 | 400 | 100 | 7650 | 41787.50 | 3600 | 900 |

---

## Data Entry Tips

1. **Use numbers, not text** — all amount and count fields must be numbers (not `"29,000.00"` as text)
2. **Leave zero values as `0`** — do not leave cells blank in amount columns
3. **Do not add extra columns** — columns after position 28 are ignored
4. **Do not merge cells** — merged cells will cause reading errors
5. **No currency symbols** — enter `29000` not `Rs. 29,000`
6. **Formulas are fine** — Excel formulas (like `=B2*0.08` for EPF) work correctly
7. **Multiple sheets** — only the first sheet is read; other sheets are ignored

---

## Common Mistakes

| Mistake | Effect | Fix |
|---------|--------|-----|
| Columns in wrong order | All amounts show as 0.00 | Reorder columns to match the table above |
| Text values in number fields | That employee's amounts show wrong | Remove currency symbols and text from number cells |
| Excel file open while generating | Error reading file | Close Excel before running the app |
| Empty row between employees | Employees after the empty row are skipped | Remove empty rows or the app stops at the gap |
| Header row missing | Row 1 employee is skipped (treated as header) | Add the header row at the top |

---

## Related Documents

- [Accountant User Guide](./02_accountant_guide.md) — how to use the app
- [Troubleshooting](./09_troubleshooting.md) — fixing data problems

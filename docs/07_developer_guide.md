# 07 · Developer Guide

> **Audience:** Developer / Software Engineer
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

This guide explains the code architecture of the Pay Slip Generator, how the files relate to each other, and how to extend or modify the system.

---

## Technology Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| GUI framework | `customtkinter` 5.2+ | Modern-looking desktop UI |
| PDF generation | `reportlab` 4.0+ | Drawing pay slips as PDF |
| Excel reading | `openpyxl` 3.1+ | Read `.xlsx` salary data |
| Image support | `pillow` | Required by customtkinter |
| Packaging | `pyinstaller` | Bundle into Windows .exe |
| Installer | Inno Setup 6 | Windows installer creator |
| Language | Python 3.10+ | All application code |

---

## File Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    USER INTERFACES                       │
│                                                         │
│  ┌─────────────────┐       ┌─────────────────────────┐ │
│  │    app.py        │       │  generate_payslips.py   │ │
│  │  (GUI — tkinter) │       │  (CLI — terminal)       │ │
│  └────────┬─────────┘       └───────────┬─────────────┘ │
│           │                             │               │
│           └─────────────┬───────────────┘               │
│                         │                               │
│              ┌───────────▼────────────┐                 │
│              │     payslip_core.py    │                 │
│              │   (Generation Engine)  │                 │
│              │                        │                 │
│              │  read_employees()      │                 │
│              │  draw_payslip()        │                 │
│              │  generate_pdf()        │                 │
│              └───────────┬────────────┘                 │
│                          │                               │
│              ┌───────────▼────────────┐                 │
│              │      openpyxl          │  reads .xlsx     │
│              │      reportlab         │  writes .pdf     │
│              └────────────────────────┘                 │
└─────────────────────────────────────────────────────────┘
```

---

## File Descriptions

### `payslip_core.py` — The Engine

This is the most important file. It contains all the PDF generation logic, completely separated from any UI code.

**Key functions:**

```python
read_employees(excel_path: str) -> list
```
- Opens the Excel file using `openpyxl`
- Reads all rows starting from row 2 (skips header)
- Skips rows where column 0 (Emp No) is empty
- Returns a list of lists (each inner list = one employee's row)

```python
draw_payslip(c, x, y, row, company_name, pay_period)
```
- Draws one complete pay slip onto a ReportLab `Canvas`
- `x`, `y` = bottom-left corner of the slip in points
- Draws: header block, employee info, 20 data rows, borders
- Uses the `COL_*` constants to access the correct column index

```python
generate_pdf(excel_path, output_path, company_name, pay_period,
             progress_callback=None, log_callback=None) -> dict
```
- Orchestrates the full generation process
- Calls `read_employees()`, calculates layout, loops through employees, calls `draw_payslip()`
- `progress_callback(pct: int)` — called with 0–100 for progress bar updates
- `log_callback(msg: str)` — called for each log message
- Returns a result dict: `{success, message, employee_count, page_count, output_path}`

**Column index constants:**
```python
COL_EMP_NO        = 0
COL_NAME          = 1
COL_DEPARTMENT    = 2
# ... (28 total)
COL_ETF3          = 27
```

These map to the 0-based column positions in the Excel sheet.

---

### `app.py` — The GUI Application

Built with `customtkinter`. Uses threads so the GUI stays responsive during generation.

**Class:** `PaySlipApp(ctk.CTk)`

**Key methods:**

| Method | Description |
|--------|-------------|
| `__init__` | Sets up window, loads saved settings, calls `_build_ui()` |
| `_build_ui()` | Creates all widgets — header, sections, buttons, log box |
| `_browse_excel()` | Opens file dialog, updates path, auto-sets output path |
| `_start_generation()` | Validates inputs, saves settings, calls `_run_in_thread()` |
| `_run_in_thread()` | Spawns a background thread so GUI doesn't freeze |
| `_on_done()` | Called from thread when generation finishes; updates UI |
| `_log()` | Thread-safe log message append to the log textbox |
| `_save_current_settings()` | Writes settings to `~/.payslip_settings.json` |

**Threading model:**

```
Main thread (GUI)
  └→ _start_generation()
       └→ _run_in_thread()  →  background thread
                                  └→ generate_pdf()  (payslip_core)
                                       └→ calls progress_callback via self.after()
                                       └→ calls log_callback via self.after()
                                  └→ self.after(0, self._on_done, result)
```

All GUI updates from the background thread use `self.after(0, ...)` to safely schedule on the main thread.

---

### `generate_payslips.py` — CLI Wrapper

A minimal script for terminal use. Imports `generate_pdf` from `payslip_core` and calls it with hardcoded settings.

**When to use:** Automation, server environments, or when the GUI cannot run (e.g. no display).

**Configuration:** Edit the constants at the top of the file:
```python
EXCEL_FILE   = "Excel/Slary_Slips.xlsx"
OUTPUT_PDF   = "PaySlips_Output.pdf"
COMPANY_NAME = "NEW LANKA CLOTHING"
PAY_PERIOD   = "May-2026"    # ← change each month
```

---

## PDF Generation — How It Works

### Coordinate System

ReportLab uses a coordinate system where `(0, 0)` is the **bottom-left** of the page.

```
  Y
  ^
  │
841│  ┌────────────────────────┐  ← Top of A4
   │  │  [Slip 1] [Slip 2]     │
   │  │                        │
   │  │  [Slip 3] [Slip 4]     │
   │  │                        │
  0│  └────────────────────────┘  ← Bottom of A4
   └──────────────────────────→ X
   0                         595
```

### Layout Calculation

```python
cols          = int((usable_w + GAP) / (SLIP_WIDTH + GAP))   # → 3
rows_per_page = int((usable_h + GAP) / (SLIP_HEIGHT + GAP))  # → 2
per_page      = cols * rows_per_page                          # → 6

# Position of each slip:
slip_x = MARGIN_LEFT + col_idx * (SLIP_WIDTH + GAP)
slip_y = PAGE_H - MARGIN_TOP - (row_idx + 1) * SLIP_HEIGHT - row_idx * GAP
```

### Draw Order for Each Slip

1. Outer rectangle border
2. Header background rectangle (dark navy)
3. Company name, "Pay Sheet", pay period text
4. Employee number, name, department, designation
5. Separator line
6. For each of 20 data rows:
   a. Background fill (colour by row type)
   b. Separator line below
   c. Label text (left)
   d. Units text (center, if applicable)
   e. Amount text (right-aligned)
7. Vertical column separator lines
8. Bottom border line

---

## Adding or Modifying Fields

### To add a new field to the pay slip

1. **Add a column constant** in `payslip_core.py`:
   ```python
   COL_NEW_FIELD = 28   # adjust index
   ```

2. **Add a row** in the `data_rows` list inside `draw_payslip()`:
   ```python
   ("New Field Label", "", _fmt(_g(row, COL_NEW_FIELD)), "earn"),
   ```
   Row styles: `"earn"`, `"gross"`, `"deduct"`, `"total_d"`, `"net"`, `"epf"`

3. **Update the Excel sheet** to include the new column

4. **Adjust `SLIP_HEIGHT`** if the new row makes the slip too tall:
   ```python
   SLIP_HEIGHT = 13.0 * cm   # increase if needed
   ```

### To change the company logo or header colour

In `payslip_core.py`, find the colour constants:
```python
C_HDR_BG  = colors.HexColor("#1a2744")   # header background
C_HDR_TEXT = colors.white                 # header text
```
Change `"#1a2744"` to any hex colour code.

### To change slip dimensions

In `payslip_core.py`:
```python
SLIP_WIDTH  = 5.0 * cm   # width of each slip
SLIP_HEIGHT = 12.0 * cm  # height of each slip
GAP         = 2 * mm      # gap between slips
```

After changing, the layout (slips per page) is recalculated automatically.

---

## Running Tests

Currently the project uses manual testing. To test:

```bash
# Test core generation
python3 generate_payslips.py
# Expected: generates PaySlips_Output.pdf with 3 sample employees

# Test imports
python3 -c "from payslip_core import generate_pdf; print('OK')"

# Test GUI starts
python3 app.py
```

---

## Dependencies Version Matrix

| Package | Min Version | Tested With |
|---------|-------------|-------------|
| Python | 3.10 | 3.13 |
| customtkinter | 5.2.0 | 5.2.2 |
| openpyxl | 3.1.0 | 3.1.5 |
| reportlab | 4.0.0 | 4.5.1 |
| pillow | 10.0.0 | 12.2.0 |
| pyinstaller | 6.0.0 | 6.x |

---

## Related Documents

- [Windows Installer Build](./06_windows_installer.md) — packaging the app
- [Data Privacy](./08_data_privacy.md) — Git and security
- [Troubleshooting](./09_troubleshooting.md) — debugging errors

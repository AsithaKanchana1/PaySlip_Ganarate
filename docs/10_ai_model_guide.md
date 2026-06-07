# AI Model Guide — PaySlip_Ganarate

> **IMPORTANT FOR AI AGENTS:** This file is the single source of truth for everything that has been built, fixed, and decided in this project. **Never delete or overwrite past session entries.** Always append a new session block at the bottom of the session log. Read this file in full before making any changes.

---

## Project Summary

| Field | Value |
|-------|-------|
| **Name** | PaySlip_Ganarate |
| **Goal** | Generate employee payslips (PDF) from a password-protected Excel salary sheet; provide a desktop GUI (customtkinter) for non-technical users |
| **Company** | New Lanka Clothing |
| **Developer** | Asitha Kanchana — https://github.com/AsithaKanchana1 |
| **Primary language** | Python 3.11+ |
| **Target platform** | Windows (installer .exe) + Linux (dev/CI) |

---

## Important Files — Quick Reference

| File | Role |
|------|------|
| `app.py` | Desktop GUI entry point (customtkinter) |
| `payslip_core.py` | **Core logic** — Excel parsing + PDF generation. Most changes happen here. |
| `generate_payslips.py` | CLI entry point (headless, no GUI) |
| `webview_launcher.py` | Webview launcher (alternative UI) |
| `index.html` | Webview HTML UI |
| `requirements.txt` | Python dependencies (runtime + dev/test) |
| `Dockerfile` | Headless Docker image for CLI use and running tests |
| `.github/workflows/build-windows.yml` | GitHub Actions: builds Windows .exe + installer + GitHub Release |
| `.github/workflows/ci.yml` | GitHub Actions: runs pytest with coverage on push/PR |
| `installer.iss` | Inno Setup script for Windows installer |
| `build_windows.bat` | Local Windows build script |
| `tests/conftest.py` | Pytest session fixtures — loads real Excel from `Excel/` folder |
| `tests/test_payslip_core.py` | 27 unit + integration tests |
| `docs/` | All documentation including this file |

---

## Excel Sheet Structure (Critical — read before touching `payslip_core.py`)

The real salary Excel (`Excel/*.xlsx`, password `1111`) has a **non-standard layout**:

### Header
- Row 0–3: empty / title rows
- **Row 4: column headers** (auto-detected by `read_employees` scanning for keyword matches)

### Column headers (39 columns, row 4)
```
Col 0:  EPF NO;          ← EPF number (manually entered, never formula)
Col 1:  No:              ← sequential employee number (manually entered, never formula)
Col 2:  Name             ← employee name
Col 3:  Net      Salary  ← cached formula result
Col 4:  PAID
Col 5:  BANK
Col 6:  Basic
Col 7:  Per day Value
Col 8:  Fixed Allowance
Col 9:  Allowance
Col 10: Food Allowance
Col 11: Att. Bonus       ← note: dots in name, normalises to 'att bonus'
Col 12: Salary Areas
Col 13: Incentive
Col 14: Normal OT (H)
Col 15: Double OT (H)
Col 16: Normal OT
Col 17: Double OT
Col 18: Incentive        ← second Incentive column → maps to COL_INCENTIVE2
Col 19: Transport Allowance
Col 20: Gross Salary
Col 21: No Pay
Col 22: No Pay           ← duplicate
Col 23: AB
Col 24: Att. Bonus       ← second Att Bonus → maps to COL_ATT_BON_DED
Col 25: Damro - Installment
Col 26: Allowance        ← second Allowance → maps to COL_ALLOW_DED
Col 27: Salary Advance
Col 28: Account Opening
Col 29: E.P.F.    8%     ← dots + spaces → normalises to 'e p f 8%'
Col 30: Late Time
Col 31: Late
Col 32: welfare
Col 33: FOOD
Col 34: FOB Deduction
Col 35: Total Deduction
Col 36: Net Salary
Col 37: E.P.F. 12%       ← normalises to 'e p f 12%'
Col 38: E.T.F.     3%    ← normalises to 'e t f 3%'
```

### Department section-header rows
Departments are **not a column** — they are embedded as special rows between employees:
- A department-header row has **only the Name cell filled in** (e.g. `'STAFF'`, `'MECHANIC'`, `'QUALITY DEPARTMENT'`)
- Columns `EPF NO;` and `No:` are **always blank** on these rows (they are manually entered, never formula-driven)
- `read_employees()` detects these rows and injects the department into `COL_DEPARTMENT` for all following employees

**Known departments in the April-2026 sheet:**
`STAFF`, `MECHANIC`, `T.MECHANIC`, `SECURITY`, `QUALITY DEPARTMENT`, `LINE RECORDER`,
`MACHINE OPERATOR`, `SUPERVISOR`, `TMO`, `HELPER (PRODUCTION)`, `IRON OPERATOR`,
`WEARE HOUSE`, `CUTTING DEPARTMENT`, `STORES HELPER`, `CLEANING DEPARTMENT`

### Formula cells
- Most salary columns (Basic, Gross, Net, EPF, etc.) are **formula-driven** in real sheets
- `openpyxl data_only=True` reads **cached formula results** — correct if the file was last saved by Microsoft Excel
- If the file was never opened/recalculated in Excel (e.g. created programmatically), formula cells return `None`
- **Do not** rely on `None` vs `0` to detect non-employee rows — formula cells can cache as `0`

---

## Key Design Decisions in `payslip_core.py`

### Column synonym mapper (`read_employees`)
- Headers are normalised: lowercased, non-alphanumeric chars → spaces, deduplicated spaces
- EPF/ETF headers with dots (e.g. `E.P.F.    8%` → `e p f 8%`) are matched with spaced forms
- Duplicate column names (Allowance ×2, Incentive ×2, Att. Bonus ×2) are mapped to primary + secondary internal columns by occurrence order
- `"id"` was removed as an EmpNo synonym to prevent `'PAID'` from false-matching

### Department detection (formula-proof)
- Uses the two manually-entered identity columns (`No:` and `EPF NO;`) as signals
- These columns are **never formula-driven** → reliably `None` on department-header rows even when other formula cells cache as `0`
- **Not** using filled-cell count (old approach broke with formula files)

### Row skipping rules (in order)
1. `if not row` — completely missing row object
2. Department-header row detected → update `current_dept`, skip (don't add as employee)
3. Name column is empty → skip (no payslip generated)

---

## How to Run Tests

```bash
# Activate venv first (Python 3.14 in .venv)
.venv/bin/pytest tests/ -v

# With coverage report
.venv/bin/pytest tests/ -v --cov=payslip_core --cov-report=term-missing

# Via Docker (runs pytest by default)
docker build -t payslip-generator .
docker run --rm payslip-generator
```

> **Note:** Tests require a real `.xlsx` file in `Excel/`. They skip automatically if none is present (e.g. fresh CI clone where `Excel/` is git-ignored). The Excel password is `1111`.

---

## How to Build the Windows EXE

```bash
# Locally on Windows:
build_windows.bat         # creates dist\PaySlipGenerator\
build_installer.bat       # creates Output\PaySlipGenerator_Setup_v1.0.exe

# Via GitHub Actions (automatic on push to main):
# → .github/workflows/build-windows.yml
# → Produces GitHub Release with installer .exe attached
```

---

## Session Log — All AI Contributions

> **Rule for AI agents:** Always ADD a new session block here. Never delete or modify past sessions. Each session must include: date, task, files changed, and test results.

---

### ✅ Session 1 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task:** Complete the "Suggested small first tasks" from the original guide

**What was done:**
- Created `tests/` directory with `__init__.py`, `conftest.py` (synthetic xlsx fixture), and `test_payslip_core.py` (27 unit tests)
- Created `.github/workflows/ci.yml` — GitHub Actions CI that runs pytest with coverage on every push/PR to `main` or `dev`
- Created `Dockerfile` — headless Python 3.11-slim image for CLI use and running tests
- Added `pytest>=7.0.0` and `pytest-cov>=4.0.0` to `requirements.txt`

| File | Action |
|------|--------|
| `tests/__init__.py` | NEW |
| `tests/conftest.py` | NEW — synthetic xlsx fixture (3 fake employees) |
| `tests/test_payslip_core.py` | NEW — 27 tests covering `_fmt`, `read_employees`, `generate_pdf` |
| `.github/workflows/ci.yml` | NEW |
| `Dockerfile` | NEW |
| `requirements.txt` | MODIFIED — added pytest deps |

**Test result:** `27 passed in 0.68s`

---

### ✅ Session 2 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task:** Switch tests to use the real Excel template (password `1111`); add `tests/fixtures/` to `.gitignore`

**What was done:**
- Rewrote `tests/conftest.py` to use the real password-protected Excel from `Excel/` folder; skips gracefully on CI
- Rewrote `tests/test_payslip_core.py` to use real file; added wrong-password test; documented openpyxl formula-cell limitation
- Deleted synthetic `tests/fixtures/` directory
- Fixed 5 real column-mapping bugs discovered by testing against real Excel:

| Bug | Fix |
|-----|-----|
| `E.P.F.    8%` normalises to `e p f 8%` (dots→spaces) — no match | Added `"e p f 8%"` / `"e p f 8"` synonyms |
| `E.P.F. 12%` normalises to `e p f 12%` | Added `"e p f 12%"` synonym |
| `E.T.F.     3%` normalises to `e t f 3%` | Added `"e t f 3%"` synonym |
| `Att. Bonus` normalises to `att bonus` — no match | Added `"att bonus"` and `"att"` synonyms |
| `PAID` falsely matched synonym `"id"` → wrong EmpNo | Removed `"id"` synonym entirely |

| File | Action |
|------|--------|
| `tests/conftest.py` | REWRITE — real Excel fixture |
| `tests/test_payslip_core.py` | REWRITE — real file + wrong-password test |
| `tests/fixtures/` | DELETED |
| `.gitignore` | MODIFIED — added `tests/fixtures/` |
| `payslip_core.py` | MODIFIED — fixed 5 synonym bugs |

**Test result:** `27 passed in 8.54s`

---

### ✅ Session 3 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task:** Fix GitHub Actions not generating the Windows EXE

**Root cause:** `build-windows.yml` had **two complete workflow definitions concatenated in one file** — GitHub Actions rejected the invalid YAML silently and never ran.

**Additional bugs fixed:**

| Bug | Fix |
|-----|-----|
| Two workflows in one YAML file | Rewrote as single clean 147-line workflow |
| `msoffcrypto-tool` and `xlrd` missing from CI pip install | Added both packages |
| `installer.iss` referenced `Excel\.gitkeep` (gitignored → doesn't exist in CI) | Removed reference; Excel dir still created via `[Dirs]` |
| `build_windows.bat` tried to `copy Excel\.gitkeep` | Replaced with `PUT_EXCEL_FILE_HERE.txt` placeholder |

| File | Action |
|------|--------|
| `.github/workflows/build-windows.yml` | REWRITE — single valid YAML workflow |
| `installer.iss` | MODIFIED — removed missing `.gitkeep` reference |
| `build_windows.bat` | MODIFIED — replaced missing `.gitkeep` copy |

**Note:** Push to GitHub requires credentials — commit is local, push manually with `git push`.

---

### ✅ Session 4 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task:** Make departments appear on payslips — they have no dedicated column

**Discovery:** The Excel sheet embeds department names as special section-header rows. When a row has only the Name cell filled (EPF and sequential No. are blank), that row is a department label for all employees that follow it.

**What was done:**
- Added department-header detection logic to `read_employees()` in `payslip_core.py`
- Scans for section-header rows (Name filled, No: blank, EPF blank)
- Tracks `current_dept` as it reads down the sheet
- Injects the department into `COL_DEPARTMENT` for every following employee row

**Result:** 16 departments detected and assigned correctly across 212 employees in the April-2026 sheet.

| File | Action |
|------|--------|
| `payslip_core.py` | MODIFIED — added `_is_dept_header()` + dept injection in parsing loop |

**Test result:** `27 passed in 10.00s`

---

### ✅ Session 5 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task 1:** Make department detection formula-proof (formulas cache as `0`, not `None`)

**Problem:** Old detection counted `non-None cells == 1`. With formula Excel files, all formula cells (Basic, Gross, EPF, etc.) cache as `0` instead of `None` → counted as filled → department rows falsely treated as employee rows.

**Fix:** Changed detection to check only the **two manually-entered identity columns** (`EPF NO;` col 0 and sequential `No:` col 1). These columns are **never formula-driven** so they stay blank on department rows regardless of what other formula cells cache.

**Task 2:** Skip rows where the Name column is empty

**Rule:** Any row with a blank Name cell → skip → no payslip generated.
This handles: blank spacer rows, summary/totals rows, formula-only placeholder rows.

**Result:** Employee count changed from 212 → 198 (14 nameless rows correctly skipped).

| File | Action |
|------|--------|
| `payslip_core.py` | MODIFIED — formula-proof `_is_dept_header()`; simplified row-skip to name-empty check |

**Test result:** `27 passed in 8.22s`

---

## Current Known Limitations

| Limitation | Status |
|-----------|--------|
| Formula cells read as `None` if file not saved by Excel after calc | Expected — documented in tests |
| 2 employees before first dept header have no department (`(none)`) | Known — they appear before first section header in the sheet |
| Designation column not present in this Excel layout | Not mapped — column doesn't exist in this sheet format |
| `openpyxl` cannot evaluate formulas — only reads cached results | By design — use Microsoft Excel to save with calculated values |

---

## Remaining Roadmap for AI Contributions

- [ ] Add `.dockerignore` to slim Docker build context
- [ ] Add data privacy checklist / redaction helpers (avoid PII in logs)
- [ ] Improve input validation and error messages in `generate_payslips.py`
- [ ] Evaluate ML-based OCR fallback for scanned salary sheets
- [ ] Add Dockerfile multi-stage build for smaller final image
- [ ] Fix the 2 employees before first department header (carry company name as fallback dept)

---

## Communication / Handoff Rules

1. **Always read this file first** before making any changes to the repo
2. **Never delete past session entries** — append new sessions at the bottom of the Session Log
3. When done with a task, add a new `### ✅ Session N` block with: date, task, files changed, test results
4. When proposing changes, include: short summary, files changed, commands to run tests, rationale
5. Keep changes minimal — maintain existing APIs and add tests for any behavioral change
6. The Excel password is `1111` — never commit real Excel files (they are gitignored)

---

### ✅ Session 6 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task:** Change layout to 4 payslips per A4 page (2×2), fill the page fully, and increase font sizes to fit the larger slip

**What was done:**

**Layout:**
- Changed from 3 cols × 2 rows (6 slips/page, with wasted space) → **2 cols × 2 rows (4 slips/page)**
- Slip size: `50mm × 120mm` → **`99mm × 142.5mm`** (nearly 2× wider, 19% taller)
- Removed hardcoded `MARGIN_LEFT/RIGHT/TOP/BOTTOM`; replaced with `PAGE_MARGIN = 5mm` (printer safe area) and `GAP = 2mm` (between slips)
- `SLIP_WIDTH` and `SLIP_HEIGHT` are now **computed from A4** minus margins and gaps: `(A4 - 2×PAGE_MARGIN - 1×GAP) / 2` per axis

**Fonts (proportional scaling):**
- `draw_payslip` now accepts `w` and `h` parameters instead of reading globals
- All font sizes computed via `scale = W / (5cm)`, clamped to sensible bounds
- Fonts approximately doubled in size (e.g. body: 5.7pt → 10pt, company: 7.5pt → 14pt)
- Column widths also proportional: label=54% of W, units=16% of W

| Font | Old | New |
|------|-----|-----|
| Company name | 7.5pt | 14pt |
| Pay Sheet / period | 6pt | 11pt |
| Employee name | 5.5pt | 11pt |
| Emp No / Dept | 5.2 / 4.8pt | 10 / 9pt |
| Body normal | 5.7pt | 10pt |
| Body bold (gross/net/total) | 6.2pt | 11pt |
| Body small (EPF/ETF) | 5.4pt | 9pt |

| File | Action |
|------|--------|
| `payslip_core.py` | MODIFIED — new layout constants, proportional `draw_payslip(w, h)`, updated `generate_pdf` positioning |

**Test result:** `27 passed in 8.42s` · PDF: 198 employees, 50 pages, 203 KB

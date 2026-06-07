## AI Model Guide for PaySlip_Ganarate

### Purpose
This document supplies AI models and contributors with a concise project overview, current state, actions taken, and recommended next steps and expectations.

### Project Summary
- **Name:** PaySlip_Ganarate
- **Goal:** Generate employee payslips (PDF) from Excel/CSV input and provide a small webview GUI for preview and launching.
- **Key entry points:** `app.py`, `generate_payslips.py`, `webview_launcher.py`.

### Important Files (quick reference)
- `app.py` — application entry for GUI/web integration.
- `generate_payslips.py` — core script to generate payslips from spreadsheets.
- `payslip_core.py` — core logic and helpers for data parsing and PDF output.
- `index.html` — webview UI used by the app.
- `requirements.txt` — Python dependencies.
- `README.md` — human-facing project readme.
- `docs/` — additional documentation and guides.

### What AI models have done (session log)

---

#### ✅ Session 1 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task: Complete all "Suggested small first tasks" from this guide**

| File | Action | Description |
|------|--------|-------------|
| `tests/__init__.py` | NEW | Python package marker |
| `tests/conftest.py` | NEW | Pytest fixtures |
| `tests/test_payslip_core.py` | NEW | 27 unit tests |
| `.github/workflows/ci.yml` | NEW | GitHub Actions CI: runs pytest on push/PR |
| `Dockerfile` | NEW | Headless Docker image (Python 3.11-slim) |
| `requirements.txt` | MODIFIED | Added `pytest>=7.0.0`, `pytest-cov>=4.0.0` |

Result: **27 passed in 0.68s** (Python 3.14.5, pytest 9.0.3)

---

#### ✅ Session 2 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task: Use real-world Excel template (Excel/ folder, password `1111`) instead of synthetic data; add tests/fixtures/ to .gitignore**

**Bugs fixed in `payslip_core.py`:**
- `'E.P.F.    8%'` → normalized to `'e p f 8%'` (dots become spaces) — added `"e p f 8%"` / `"e p f 8"` synonyms
- `'E.P.F. 12%'` → normalized to `'e p f 12%'` — added `"e p f 12%"` synonym
- `'E.T.F.     3%'` → normalized to `'e t f 3%'` — added `"e t f 3%"` synonym
- `'Att. Bonus'` → normalized to `'att bonus'` — added `"att bonus"` and `"att"` synonyms
- `'PAID'` → was falsely matching synonym `"id"` → removed `"id"` as an EmpNo synonym to prevent substring false-positives

**Note discovered:** The real Excel template uses formula-driven salary cells. `openpyxl data_only=True` reads cached formula results — if the file hasn't been recalculated and saved by Excel, those cells come back as `None`. This is expected behaviour and is documented in `test_basic_salary_or_formula_column_present`.

| File | Action | Description |
|------|--------|-------------|
| `tests/conftest.py` | REWRITE | Uses real Excel from `Excel/` folder (password `1111`); skips gracefully on CI |
| `tests/test_payslip_core.py` | REWRITE | Tests use real file; added wrong-password test; formula-cell test documents openpyxl limitation |
| `tests/fixtures/` | DELETED | Removed synthetic fixture directory (no longer needed) |
| `.gitignore` | MODIFIED | Added `tests/fixtures/` section |
| `payslip_core.py` | MODIFIED | Fixed 5 synonym-mapping bugs for real-world column headers |

Result: **27 passed in 8.54s** (Python 3.14.5, pytest 9.0.3)

---

#### ✅ Session 3 — 2026-06-07 (Antigravity / Claude Sonnet)

**Task: Fix GitHub Actions not generating the Windows EXE**

**Root cause:** `.github/workflows/build-windows.yml` contained **two full workflow definitions concatenated into one file** — GitHub Actions rejects this as invalid YAML and never runs the build.

**Additional bugs found:**
- `installer.iss` referenced `Excel\.gitkeep` which is gitignored → Inno Setup would fail with "file not found"
- `build_windows.bat` also tried to `copy Excel\.gitkeep` → silent failure on CI
- `msoffcrypto-tool` and `xlrd` were missing from the CI pip install command

| File | Action | Description |
|------|--------|-------------|
| `.github/workflows/build-windows.yml` | REWRITE | Single clean workflow (was two workflows concatenated); added `msoffcrypto-tool`/`xlrd` to install; fixed PowerShell release step to reliably glob the output exe |
| `installer.iss` | MODIFIED | Removed `Excel\.gitkeep` source reference; Excel dir still created via `[Dirs]` section |
| `build_windows.bat` | MODIFIED | Replaced `copy Excel\.gitkeep` with creating a `PUT_EXCEL_FILE_HERE.txt` placeholder |

- `TestReadEmployees` (9 tests) — Excel parsing, column mapping, error handling
- `TestGeneratePdf` (9 tests) — PDF generation: file creation, employee count, PDF magic bytes, callbacks

**Commands to run tests:**
```bash
# Using venv (Python 3.14):
.venv/bin/pytest tests/test_payslip_core.py -v

# Or with coverage:
.venv/bin/pytest tests/ -v --cov=payslip_core --cov-report=term-missing

# Or via Docker:
docker build -t payslip-generator .
docker run --rm payslip-generator
```

---

### How AI models should use this repo
- **Context window strategy:** Provide only the relevant files (module under change, tests, and small supporting files). Prefer focused snippets over whole files when possible.
- **Task framing:** Use explicit instructions: goal, constraints, input examples, and expected output format (diff patch, new file, unit test, or step-by-step plan).
- **Safety & privacy constraints:** Treat `Excel/` data and any sample data as potentially sensitive; do not print real personal data. Prefer synthetic examples.
- **Coding style:** Keep changes minimal and maintain existing APIs. Add tests for behavioral changes.

### Recommended prompt templates for common tasks
- Bugfix: "Find and fix the bug in `payslip_core.py` that causes X when given Y; produce a concise patch and a unit test." 
- Feature: "Add feature: [short description]. Update `generate_payslips.py` to accept [new flag], modify core logic, and add tests." 
- Refactor: "Refactor `generate_payslips.py` to extract reusable functions. Keep behavior and CLI stable. Provide unit tests." 
- Tests: "Write pytest unit tests for `payslip_core.py` covering parsing and PDF rendering using synthetic input." 

### Input / Output expectations for automated agents
- **Inputs provided to the model:** file path(s), small code snippets (<=500 lines), failing test output, sample input spreadsheet (synthetic), and exact reproduction steps.
- **Desired outputs:** an apply_patch-style diff, or a clear set of file changes with tests. When returning code, include a short explanation and commands to run tests.

### Quick run & dev notes
- To install deps: `pip install -r requirements.txt`.
- Typical run: `python generate_payslips.py --input Excel/sample.xlsx --out out/` (adapt per actual CLI in repo).
- Environment: Linux (developer machine), user prefers Fish shell; CI should use POSIX-compatible shells.
- Tests: `.venv/bin/pytest tests/ -v` (all 27 tests pass as of 2026-06-07)

### Validation & testing
- pytest is now in `requirements.txt` as a dev dependency.
- `tests/` directory exists with synthetic Excel/CSV fixtures and tests that assert output PDF existence and basic content checks.
- CI runs automatically on push to `main` / `dev` branches via `.github/workflows/ci.yml`.

### Remaining roadmap for AI contributions
- Add data privacy checklist and redaction helpers to avoid leaking real PII in outputs.
- Evaluate adding an ML-based OCR fallback for scanned input spreadsheets or images.
- Improve input validation and error messages in `generate_payslips.py`.
- Add `.dockerignore` to slim down Docker build context.
- Add a `Dockerfile` multi-stage build for even smaller final image.

### Communication / Handoff
- When proposing changes, always include: 1) short summary, 2) files changed, 3) commands to run tests, 4) rationale for design choices.
- When done, update the **Session log** table above with date, agent name, files changed, and test results.

---

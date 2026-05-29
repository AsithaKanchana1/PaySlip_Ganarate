# 08 · Data Privacy & Git

> **Audience:** Developer / IT Administrator
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

The Pay Slip Generator handles **sensitive personal salary data**. This document explains what data is kept private, how the `.gitignore` works, and best practices for data security.

---

## What Data is Sensitive

| File / Data | Sensitivity | Reason |
|-------------|-------------|--------|
| `Excel/Slary_Slips.xlsx` | 🔴 HIGH | Contains salary, EPF, ETF, deductions for all employees |
| `PaySlips_Output.pdf` | 🔴 HIGH | Contains same data in printable form |
| `~/.payslip_settings.json` | 🟡 MEDIUM | Contains file paths used on this machine |
| `app.py`, `payslip_core.py` | 🟢 LOW | Application code only — no personal data |
| `README.md`, `docs/` | 🟢 LOW | Documentation only |

---

## The `.gitignore` File

The `.gitignore` file in the project root tells Git which files to **never track or upload to GitHub**.

### Current rules:

```gitignore
# Excel Data Files (private salary data)
Excel/*.xlsx
Excel/*.xls
Excel/*.csv

# Generated PDF Output
PaySlips_Output.pdf
*.pdf

# Windows Build Output
dist/
build/
Output/
*.spec
__pycache__/

# Virtual Environments
venv/
.venv/

# OS / Editor
.DS_Store
Thumbs.db
.idea/
.vscode/
```

### What this means in practice

| What you do | What happens |
|-------------|-------------|
| `git add .` | Excel and PDF files are automatically excluded |
| `git push` | Only code files reach GitHub |
| Someone clones the repo | They get the code but NOT the salary data |
| You add a new `.xlsx` file | Still excluded automatically |

---

## Verifying Nothing Private was Committed

Check which files Git is currently tracking:

```bash
git ls-files
```

You should NOT see any `.xlsx`, `.xls`, or `.pdf` files in the output.

If you do see them, remove them immediately (see below).

---

## How to Remove a File That Was Accidentally Committed

If a salary file was committed to Git, follow these steps:

### Remove from tracking (keeps file on disk)

```bash
# Remove the specific file from Git
git rm --cached Excel/Slary_Slips.xlsx
git rm --cached PaySlips_Output.pdf

# Commit the removal
git commit -m "Remove private data files from Git tracking"

# Push to GitHub
git push
```

> ⚠️ This removes the file from **future** commits but the file may still exist in Git **history**.

### Completely remove from Git history (nuclear option)

If the file contains truly sensitive data and was pushed to GitHub:

```bash
# Install git-filter-repo (safer than git filter-branch)
pip install git-filter-repo

# Remove the file from ALL history
git filter-repo --path Excel/Slary_Slips.xlsx --invert-paths

# Force push to overwrite GitHub history
git push origin --force --all
```

> ⚠️ **Force-pushing rewrites history.** Everyone who has cloned the repo must re-clone. Coordinate with all team members before doing this.

> 💡 After removing from history, also **revoke and rotate** any secrets if API keys or passwords were accidentally committed.

---

## The `Excel/.gitkeep` File

The `Excel/` directory contains a special file called `.gitkeep`:

```
Excel/
├── .gitkeep        ← tracked by Git (keeps folder structure on GitHub)
└── Slary_Slips.xlsx  ← NOT tracked (excluded by .gitignore)
```

**Why this exists:** Git does not track empty folders. Without `.gitkeep`, the `Excel/` folder would not appear in the repository, and users who clone the repo would not know where to put their Excel file.

The `.gitkeep` file itself contains only a comment explaining its purpose — no private data.

---

## Setting Up on a New Machine

When someone clones the repo, they must manually add the Excel file:

```bash
git clone https://github.com/AsithaKanchana1/PaySlip_Ganarate.git
cd PaySlip_Ganarate

# Copy your Excel file into the Excel/ folder
cp /path/to/Slary_Slips.xlsx Excel/

# Run the app
python3 app.py
```

**The Excel file is NEVER shared through Git.** It must be copied manually or shared through a secure internal channel (e.g. company shared drive, encrypted USB).

---

## Data Storage Summary

| Location | What is stored | Who has access |
|----------|---------------|----------------|
| GitHub repository | Code only (no salary data) | Public |
| Developer's computer | Code + Excel + PDF | Developer only |
| Accountant's computer | Installed app + Excel + PDF | Accountant only |
| `~/.payslip_settings.json` | File paths only (no salary data) | Current user only |

---

## Best Practices

1. **Never email the Excel file** — use the company shared drive or encrypted transfer
2. **Keep the PDF in a secure folder** — it contains everyone's salary
3. **Delete old PDFs** after distributing slips — or move to a secure archive folder
4. **Do not take screenshots** of pay slips and share them publicly
5. **Lock your screen** when leaving your desk with the app open
6. **Use Windows BitLocker** or equivalent to encrypt the drive storing salary data
7. **Check `git status`** before every `git push` to confirm no private files are staged

---

## Related Documents

- [Developer Guide](./07_developer_guide.md) — code architecture
- [Windows Installer Build](./06_windows_installer.md) — distribution
- [Getting Started](./01_getting_started.md) — initial setup on new machines

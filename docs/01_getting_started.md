# 01 · Getting Started

> **Audience:** Everyone — first-time setup on any machine
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

This guide walks you through getting the Pay Slip Generator running for the first time, whether you are:
- An **accountant** installing the ready-made Windows app
- A **developer** setting up the Python project

---

## Option A — Windows Installation (Accountant / End User)

This is the simplest path. No Python or technical knowledge needed.

### What you need
- A Windows 10 or Windows 11 computer
- The `PaySlipGenerator_Setup.exe` file (provided by your IT person)

### Steps

**Step 1 — Run the installer**

Double-click `PaySlipGenerator_Setup.exe`.

```
Windows may show a SmartScreen warning:
  "Windows protected your PC"

Click: More info → Run anyway
(This is normal for software not yet published to the Microsoft Store)
```

**Step 2 — Follow the setup wizard**

| Screen | Action |
|--------|--------|
| Welcome | Click **Next** |
| Select destination | Leave default, click **Next** |
| Select tasks | ✅ Keep "Create Start Menu shortcut" checked |
| Desktop shortcut | Optional — tick if you want one |
| Ready to install | Click **Install** |
| Finish | Click **Finish** |

**Step 3 — Note the installation folder**

After installation, a message will appear showing where to put your Excel file:

```
C:\Program Files\NewLankaClothing\PaySlipGenerator\Excel\
```

> ⚠️ You must copy your `Slary_Slips.xlsx` file into this `Excel\` folder before using the app.

**Step 4 — Launch the app**

Open **Start Menu** → search for **Pay Slip Generator** → click to open.

✅ You are ready to use the app. See [Accountant User Guide](./02_accountant_guide.md) for monthly usage.

---

## Option B — Python Setup (Developer / Linux / macOS)

### What you need
- Python 3.10 or higher
- pip (comes with Python)
- Git (optional, for cloning)

### Check your Python version

```bash
python3 --version
# Should show: Python 3.10.x or higher
```

If Python is not installed, download it from [python.org](https://www.python.org/downloads/).

### Step 1 — Get the project

**Via Git (recommended):**
```bash
git clone https://github.com/AsithaKanchana1/PaySlip_Ganarate.git
cd PaySlip_Ganarate
```

**Or download the ZIP** from GitHub and extract it.

### Step 2 — Install dependencies

```bash
pip install -r requirements.txt
```

On **Arch Linux** (externally managed Python):
```bash
pip install -r requirements.txt --break-system-packages
```

This installs:
- `customtkinter` — modern GUI framework
- `openpyxl` — read Excel files
- `reportlab` — generate PDF files
- `pillow` — image support

### Step 3 — Add your Excel file

Copy your salary Excel file into the `Excel/` folder:

```
Pay-Slip_Genarate/
└── Excel/
    └── Slary_Slips.xlsx   ← copy here
```

### Step 4 — Run the app

```bash
# Graphical desktop app (recommended)
python3 app.py

# OR command-line only
python3 generate_payslips.py
```

---

## Verifying the Installation

After launching, you should see:

```
┌─────────────────────────────────────────┐
│  🧾  Pay Slip Generator                 │
│       New Lanka Clothing                │
├─────────────────────────────────────────┤
│  ⚙️ Company Settings                   │
│  ...                                    │
└─────────────────────────────────────────┘
```

If anything goes wrong, see [Troubleshooting](./09_troubleshooting.md).

---

## Next Steps

- 👩‍💼 **Accountant:** Read [Accountant User Guide](./02_accountant_guide.md)
- 💻 **Developer:** Read [Developer Guide](./07_developer_guide.md)
- 📊 **Setting up Excel:** Read [Excel File Format](./03_excel_format.md)

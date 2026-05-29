# 06 · Windows Installer Build Guide

> **Audience:** Developer / IT Administrator
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

This guide explains how to package the Pay Slip Generator into a professional Windows installer (`.exe`) that the accountant can install like any normal Windows software.

The process has two stages:
1. **PyInstaller** — bundles Python + all libraries into a standalone `.exe`
2. **Inno Setup** — wraps the `.exe` into an installer with Start Menu shortcut and uninstaller

---

## Prerequisites

Do this on a **Windows 10 or 11** machine:

| Software | Version | Download |
|----------|---------|----------|
| Python | 3.10 or higher | [python.org](https://www.python.org/downloads/) |
| Git | Any | [git-scm.com](https://git-scm.com/) |
| Inno Setup | 6.x | [jrsoftware.org/isdl.php](https://jrsoftware.org/isdl.php) |

> ⚠️ When installing Python, **tick "Add Python to PATH"** on the first screen of the installer.

---

## Step 1 — Get the Project

Open **Command Prompt** (search `cmd` in Start Menu):

```cmd
git clone https://github.com/AsithaKanchana1/PaySlip_Ganarate.git
cd PaySlip_Ganarate
```

Or download and extract the ZIP from GitHub.

---

## Step 2 — Run the Build Script

Double-click `build_windows.bat` — or run in Command Prompt:

```cmd
build_windows.bat
```

This script does the following automatically:

| Step | What happens |
|------|--------------|
| 1 | Installs `customtkinter`, `openpyxl`, `reportlab`, `pillow`, `pyinstaller` via pip |
| 2 | Cleans any previous build output |
| 3 | Runs PyInstaller to bundle everything into a folder |
| 4 | Copies README and creates the Excel folder inside the output |

**Expected output** (takes 1–2 minutes):
```
[1/4] Installing Python dependencies... Done.
[2/4] Cleaning previous build... Done.
[3/4] Building standalone .exe (this takes 1-2 minutes)...
      ... (PyInstaller output) ...
      Done.
[4/4] Packaging... Done.

=====================================================
  BUILD COMPLETE!
=====================================================
  Your application is in:
  dist\PaySlipGenerator\

  To run:  dist\PaySlipGenerator\PaySlipGenerator.exe
```

After this step, you can already run the app directly:
```cmd
dist\PaySlipGenerator\PaySlipGenerator.exe
```

---

## Step 3 — Add an Icon (Optional but Recommended)

If you have a custom icon file (`icon.ico`), place it in the project root before building:
```
Pay-Slip_Genarate/
└── icon.ico     ← place here (256x256 pixels recommended)
```

The build script and installer script both reference `icon.ico` automatically.

> 💡 To convert a PNG to ICO format, use [icoconvert.com](https://icoconvert.com/) or any image editor.

---

## Step 4 — Create the Windows Installer

1. Open **Inno Setup** (from Start Menu)
2. Click **File → Open** → browse to `installer.iss` in the project folder
3. Press **F9** (or click **Build → Compile**)
4. Wait for compilation (about 30 seconds)

**Output file:**
```
Pay-Slip_Genarate/
└── Output/
    └── PaySlipGenerator_Setup_v1.0.exe   ← this is the installer
```

---

## Step 5 — Test the Installer

Before distributing, test the installer on a clean Windows machine (or a virtual machine):

1. Copy `PaySlipGenerator_Setup_v1.0.exe` to the test machine
2. Run the installer and follow the wizard
3. Launch the app from Start Menu
4. Copy a test Excel file and generate a PDF
5. Verify the PDF is correct

---

## Step 6 — Distribute

Share `PaySlipGenerator_Setup_v1.0.exe` with the accountant via:
- USB drive
- Shared network folder
- Email (if file size allows)
- GitHub Releases page

The accountant only needs this one file.

---

## What the Installer Creates

When the accountant runs the installer, it creates:

```
C:\Program Files\NewLankaClothing\PaySlipGenerator\
├── PaySlipGenerator.exe     ← main application
├── Excel\                   ← empty folder for salary data
├── README.md                ← instructions
├── _internal\               ← Python libraries (do not touch)
└── ...
```

Start Menu entries:
```
Start Menu → New Lanka Clothing →
    ├── Pay Slip Generator
    └── Uninstall Pay Slip Generator
```

---

## Updating the App (New Version)

When you make changes to the code:

1. Make your code changes
2. Update `#define AppVersion` in `installer.iss` (e.g. change `"1.0"` to `"1.1"`)
3. Run `build_windows.bat` again
4. Compile `installer.iss` again in Inno Setup
5. Distribute the new `PaySlipGenerator_Setup_v1.1.exe`

The new installer will automatically uninstall the old version and install the new one.

---

## Troubleshooting the Build

| Problem | Cause | Fix |
|---------|-------|-----|
| `pip` not found | Python not in PATH | Reinstall Python, tick "Add to PATH" |
| `pyinstaller` not found | pip install failed | Run `pip install pyinstaller` manually |
| Build fails with import error | Missing library | Run `pip install -r requirements.txt` first |
| `icon.ico` not found warning | Icon file missing | Either add `icon.ico` or remove `--icon` flag from the bat file |
| Inno Setup compile error | `dist\` folder missing | Run `build_windows.bat` first |
| Antivirus blocks the exe | False positive | Add an exception in antivirus; sign the exe with a code signing certificate for production |

---

## Related Documents

- [Getting Started](./01_getting_started.md) — initial setup
- [Developer Guide](./07_developer_guide.md) — code architecture
- [Data Privacy](./08_data_privacy.md) — what to keep private

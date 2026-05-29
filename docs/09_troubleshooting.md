# 09 · Troubleshooting

> **Audience:** Everyone
> **Back to:** [Documentation Index](./doc.md)

---

## Overview

This guide covers the most common problems users encounter with the Pay Slip Generator, organised by category.

---

## App Won't Start

### ❌ "Windows protected your PC" (SmartScreen warning)

**Cause:** Windows SmartScreen blocks newly-created .exe files that aren't code-signed.

**Fix:**
1. Click **More info**
2. Click **Run anyway**

This is normal for internal/custom software. It does not mean the app is dangerous.

---

### ❌ App opens and immediately closes

**Cause:** Usually a Python error on startup.

**Fix (Windows):**
1. Open Command Prompt
2. Navigate to the app folder:
   ```cmd
   cd "C:\Program Files\NewLankaClothing\PaySlipGenerator"
   PaySlipGenerator.exe
   ```
3. Read the error message that appears before it closes

**Fix (Python / Developer):**
```bash
python3 app.py
# Read the error in the terminal
```

---

### ❌ `ModuleNotFoundError: No module named 'customtkinter'`

**Cause:** Dependencies not installed.

**Fix:**
```bash
pip install -r requirements.txt
```

On Arch Linux:
```bash
pip install -r requirements.txt --break-system-packages
```

---

### ❌ `ImportError: libtk8.6.so: cannot open shared object file`

**Cause:** Tkinter system library missing (Linux only).

**Fix:**
- Ubuntu/Debian: `sudo apt install python3-tk`
- Arch Linux: `sudo pacman -S tk`
- Fedora: `sudo dnf install python3-tkinter`

---

## Excel File Problems

### ❌ "Excel file not found"

**Cause:** The file path in the app points to a file that no longer exists.

**Fix:**
1. Click **📁 Browse** again
2. Navigate to where your Excel file is currently saved
3. Select it and click **Open**

---

### ❌ "PermissionError" when reading Excel

**Cause:** The Excel file is currently open in Microsoft Excel.

**Fix:**
1. Switch to Microsoft Excel
2. Close the file (Ctrl+W or File → Close)
3. Try generating again

> ⚠️ You must close the file in Excel before the app can read it.

---

### ❌ All amounts show as `0.00`

**Cause:** The Excel columns are in the wrong order. The app reads columns by position, not by header name.

**Fix:**
1. Open your Excel file
2. Compare the column order to the table in [Excel File Format](./03_excel_format.md)
3. Rearrange columns to match the required order
4. Save and try again

---

### ❌ Some employees are missing from the PDF

**Cause 1:** There is an empty row between employees.
**Fix:** Remove the empty row in Excel. The app stops reading when it finds an empty `Emp No` cell.

**Cause 2:** Those employees' rows have no value in the first column (Emp No).
**Fix:** Make sure every employee has an Emp No in column A.

---

### ❌ Employee name or department is missing/blank on the slip

**Cause:** The Name or Department cell is empty in Excel.
**Fix:** Fill in the missing values in the Excel file and regenerate.

---

### ❌ Numbers appear as text (e.g. `"29,000"` with quotes or commas)

**Cause:** The cell is formatted as Text or contains currency symbols/commas.
**Fix:**
1. Select the cell in Excel
2. Remove any `Rs.`, `$`, `,` characters
3. Change the cell format to **Number**
4. Save and regenerate

---

## PDF Generation Problems

### ❌ "PermissionError" when saving PDF

**Cause:** The output PDF is open in a PDF viewer.

**Fix:**
1. Close the PDF in your viewer (Adobe Reader, Edge, etc.)
2. Try generating again

---

### ❌ PDF is generated but the file is empty or corrupted

**Cause:** The generation process was interrupted, or disk is full.

**Fix:**
1. Check available disk space
2. Try saving the PDF to a different location
3. Delete the old PDF file manually, then regenerate

---

### ❌ Wrong month or year on the pay slips

**Cause:** The Pay Period was set incorrectly.

**Fix:**
1. In the app, correct the Month dropdown and Year field
2. Click **Generate** again
3. The PDF will be regenerated with the correct period

---

## Printing Problems

### ❌ Slips are the wrong size when printed

**Cause:** The PDF was printed with "Fit to page" or a scale other than 100%.

**Fix:**
1. Open the PDF
2. Press Ctrl+P
3. Find the size/scale setting
4. Set it to **Actual size** or **100%**
5. Make sure "Fit to page" is **unchecked**
6. Print again

---

### ❌ Slips are cut off at the edges

**Cause:** Printer margins are too large.

**Fix:**
1. In print settings, look for **Margins** or **Borderless printing**
2. Set margins to **Minimum** or **None**
3. Some printers cannot print to the very edge — ensure the slip content (not margin) is within the printable area

---

### ❌ Some slips on the last page are missing

**Cause:** The last page does not have enough employees to fill 6 slots. This is normal.

**Explanation:** If you have 14 employees, the last page will have only 2 slips. The remaining 4 spaces are empty. This is correct behaviour.

---

## Windows Installer Problems

### ❌ Build fails: "pyinstaller is not recognized"

**Cause:** PyInstaller was not installed, or Python is not in PATH.

**Fix:**
```cmd
pip install pyinstaller
```
If that fails: reinstall Python and check "Add Python to PATH".

---

### ❌ Antivirus deletes the `.exe` after building

**Cause:** Antivirus software flags freshly-built PyInstaller executables as suspicious (false positive).

**Fix:**
1. Add the `dist\` folder to your antivirus exclusion list
2. For production distribution, consider code-signing the `.exe` with a certificate

---

### ❌ Inno Setup says `dist\PaySlipGenerator` not found

**Cause:** `build_windows.bat` was not run first, or the build failed.

**Fix:**
1. Run `build_windows.bat` first
2. Check for errors in the output
3. Once `dist\PaySlipGenerator\PaySlipGenerator.exe` exists, compile the installer

---

## Getting More Help

1. **Check the Activity Log** in the app — it shows detailed error messages
2. **Run from terminal** to see full Python error output:
   ```bash
   python3 app.py
   ```
3. **Check the GitHub repository** for known issues and updates:
   [github.com/AsithaKanchana1/PaySlip_Ganarate](https://github.com/AsithaKanchana1/PaySlip_Ganarate)

---

## Related Documents

- [Getting Started](./01_getting_started.md) — initial installation
- [Excel File Format](./03_excel_format.md) — correct column order
- [PDF Output & Printing](./05_pdf_output.md) — print settings
- [Developer Guide](./07_developer_guide.md) — code-level debugging

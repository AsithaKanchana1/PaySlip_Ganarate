@echo off
:: ============================================================
::  build_windows.bat
::  Run this on a Windows PC to create the installer .exe
::  Requirements: Python 3.10+ installed on Windows
:: ============================================================
title Building Pay Slip Generator...
setlocal EnableDelayedExpansion

echo.
echo  =====================================================
echo   New Lanka Clothing - Pay Slip Generator Builder
echo  =====================================================
echo.

:: Step 1 — Install dependencies
echo [1/4] Installing Python dependencies...
pip install customtkinter openpyxl reportlab pillow pyinstaller --quiet
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    pause & exit /b 1
)
echo        Done.
echo.

:: Step 2 — Clean old build
echo [2/4] Cleaning previous build...
if exist dist\PaySlipGenerator rmdir /s /q dist\PaySlipGenerator
if exist build rmdir /s /q build
echo        Done.
echo.

:: Step 3 — Build the exe with PyInstaller
echo [3/4] Building standalone .exe (this takes 1-2 minutes)...

:: Check if icon.ico exists — make it optional
set ICON_FLAG=
if exist icon.ico (
    set ICON_FLAG=--icon "icon.ico"
    echo        Found icon.ico — will use custom icon.
) else (
    echo        icon.ico not found — building without custom icon.
)

pyinstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "PaySlipGenerator" ^
    --add-data "payslip_core.py;." ^
    --hidden-import customtkinter ^
    --hidden-import openpyxl ^
    --hidden-import reportlab ^
    --hidden-import PIL ^
    --collect-all customtkinter ^
    --collect-data customtkinter ^
    %ICON_FLAG% ^
    app.py

if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed.
    pause & exit /b 1
)
echo        Done.
echo.

:: Step 4 — Copy README into dist folder
echo [4/4] Packaging...
copy README.md dist\PaySlipGenerator\README.md >nul 2>&1
:: Create an Excel subfolder so users know where to put their file
mkdir dist\PaySlipGenerator\Excel >nul 2>&1
echo Place your salary Excel file here. > dist\PaySlipGenerator\Excel\PUT_EXCEL_FILE_HERE.txt

echo.
echo  =====================================================
echo   BUILD COMPLETE!
echo  =====================================================
echo.
echo   Your application is in:
echo   dist\PaySlipGenerator\
echo.
echo   To run:  dist\PaySlipGenerator\PaySlipGenerator.exe
echo.
echo   To make an installer, run:  build_installer.bat
echo  =====================================================
echo.
pause

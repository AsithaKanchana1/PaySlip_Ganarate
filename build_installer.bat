@echo off
:: build_installer.bat
:: Run this on Windows after building the PyInstaller output to create the Inno Setup installer

echo.
echo =====================================================
echo  Building Inno Setup installer for PaySlipGenerator
echo =====================================================

:: Check for ISCC in PATH
where ISCC.exe >nul 2>nul
if %errorlevel% neq 0 (
    echo ISCC.exe not found in PATH. Please install Inno Setup (https://jrsoftware.org/isdl.php)
    echo Or install via Chocolatey: choco install innosetup -y
    pause & exit /b 1
)

if not exist installer.iss (
    echo installer.iss not found in current directory.
    pause & exit /b 1
)

echo Compiling installer...
ISCC.exe installer.iss
if %errorlevel% neq 0 (
    echo ERROR: Inno Setup compilation failed.
    pause & exit /b 1
)

echo Installer build complete. Check the Output\ folder for the Setup executable.
pause

@echo off
:: ─────────────────────────────────────────────────────────────────────────────
:: start.bat  —  Kumulus Partners PDF Tools
:: Lives in watermarking/  —  code lives in watermarking/code/
:: Double-click this file to launch the GUI.
:: ─────────────────────────────────────────────────────────────────────────────

cd /d "%~dp0"

:: ── Check for Python ─────────────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% == 0 goto install_packages

echo.
echo  Python not found. Installing via winget...
echo.
winget install --id Python.Python.3.11 --silent --accept-source-agreements --accept-package-agreements

if %errorlevel% neq 0 (
    echo.
    echo  Automatic install failed.
    echo  Download manually from: https://www.python.org/downloads/
    echo  During installation, check:  [x] Add Python to PATH
    echo.
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo  Python installed. Please close this window and run start.bat again.
pause
exit /b 0

:: ── Install / verify packages ────────────────────────────────────────────────
:install_packages
python -m pip install --quiet pypdf reportlab openpyxl fonttools brotli cryptography tkinterdnd2 python-docx pywin32 pillow

:: pywin32 needs a post-install step to register its DLLs (pythoncom,
:: pywintypes). Without it, 'import win32com' fails even after pip succeeds.
python -m pywin32_postinstall -install >nul 2>&1

:: ── Launch ───────────────────────────────────────────────────────────────────
python code\launcher.py

if %errorlevel% neq 0 (
    echo.
    echo  The app exited with an error. See message above.
    pause
)

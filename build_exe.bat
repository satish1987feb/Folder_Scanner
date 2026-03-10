@echo off
echo ============================================
echo   Building Folder Scanner Desktop (.exe)
echo ============================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install pandas openpyxl pyinstaller --quiet

echo [2/3] Building executable...
pyinstaller --onefile --windowed --name "FolderScanner" --clean desktop_app.py

echo [3/3] Done!
echo.
echo Your .exe is at:  dist\FolderScanner.exe
echo Share this file with anyone — no Python needed.
echo.
pause

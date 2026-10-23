@echo off
chcp 65001 >nul 2>&1
title Build DWG2PDF.exe
cd /d "%~dp0"

set "PY=C:\Users\朱柳军\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" (
    where python >nul 2>&1
    if errorlevel 1 goto nopython
    set "PY=python"
)

echo Building single-file exe ...
"%PY%" -m PyInstaller --noconfirm --onefile --noconsole --name "DWG批量转PDF" ^
    --icon "app.ico" --add-data "ui.html;." ^
    --distpath "dist" --workpath "build" --specpath "." dwg2pdf_ui.py

if errorlevel 1 goto failed
echo.
echo Done: dist\DWG批量转PDF.exe
echo Copy that file anywhere you like; it does not need Python.
pause
exit /b 0

:failed
echo.
echo Build failed. See messages above.
pause
exit /b 1

:nopython
echo Python not found. Install Python 3.9+ and run: pip install pywin32 pyinstaller
pause
exit /b 1

@echo off
chcp 65001 >nul 2>&1
title DWG to PDF
cd /d "%~dp0"

set "PY=C:\Users\朱柳军\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" (
    where python >nul 2>&1
    if errorlevel 1 goto nopython
    set "PY=python"
)

"%PY%" "%~dp0dwg2pdf_ui.py" %*
if errorlevel 1 (
    echo.
    echo Something went wrong. See message above.
    pause
)
exit /b

:nopython
echo.
echo Python not found. Please install Python 3.9+ and pywin32 first.
echo   pip install pywin32
echo.
pause
exit /b 1

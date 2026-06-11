@echo off
chcp 65001 >nul
title HAY Sale Report - Fetch from Odoo

echo ============================================
echo   HAY Sale Report - Odoo Direct Fetch
echo ============================================
echo.

set PYTHON=

REM Try py launcher first (most reliable on Windows)
py --version >nul 2>&1
if %ERRORLEVEL% == 0 (
    set PYTHON=py
    goto :found
)

REM Try python
python --version >nul 2>&1
if %ERRORLEVEL% == 0 (
    set PYTHON=python
    goto :found
)

REM Try common install paths
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist %%P (
        set PYTHON=%%P
        goto :found
    )
)

echo ERROR: Python not found. Please install Python from https://python.org
pause
exit /b 1

:found
echo Using Python: %PYTHON%
echo.

cd /d "%~dp0"

REM v3.10.1: Force UTF-8 for Python stdout/stderr (avoid cp1252 crash on Unicode)
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

REM Check and install pywin32 if missing (required for PDF export)
%PYTHON% -c "import win32com.client" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Installing pywin32...
    %PYTHON% -m pip install pywin32 -q
    if %ERRORLEVEL% neq 0 (
        echo WARNING: Could not install pywin32 - PDF step will be skipped
    ) else (
        echo pywin32 installed OK
    )
    echo.
)

REM v3.6: Check and install requests + python-dotenv (required for email step)
%PYTHON% -c "import requests, dotenv" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Installing requests + python-dotenv...
    %PYTHON% -m pip install requests python-dotenv -q
    if %ERRORLEVEL% neq 0 (
        echo WARNING: Could not install email deps - email step will be skipped
    ) else (
        echo requests + python-dotenv installed OK
    )
    echo.
)

REM v3.7: Check and install tkcalendar (required for date picker dialog)
%PYTHON% -c "import tkcalendar" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Installing tkcalendar...
    %PYTHON% -m pip install tkcalendar -q
    if %ERRORLEVEL% neq 0 (
        echo WARNING: Could not install tkcalendar - date picker will be skipped ^(runs for today^)
    ) else (
        echo tkcalendar installed OK
    )
    echo.
)

REM Close any leftover Excel process from a previous failed run
taskkill /F /IM EXCEL.EXE >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Closed leftover Excel process
    timeout /t 2 /nobreak >nul
)

echo Launching date picker...
echo.
REM v3.7: If tkcalendar is available, show the calendar dialog (30s auto-run).
REM Otherwise, fall back to running daily_report.py directly for today.
%PYTHON% -c "import tkcalendar" >nul 2>&1
if %ERRORLEVEL% == 0 (
    %PYTHON% run_for_date.py
) else (
    echo WARNING: tkcalendar not available - running for today's date
    %PYTHON% daily_report.py --fetch-odoo
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: Something went wrong - see error above
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Done!
echo ============================================
echo.
pause

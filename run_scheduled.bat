@echo off
setlocal EnableExtensions
REM ============================================================================
REM run_scheduled.bat - invoked by Windows Task Scheduler at 6 AM Bangkok time
REM Differences from run_fetch_odoo.bat:
REM   1. No date picker (defaults to YESTERDAY via v3.9)
REM   2. No `pause` at the end (would block scheduled execution forever)
REM   3. All output captured to _Archive\scheduled_logs\<timestamp>.log
REM ============================================================================

cd /d "%~dp0"

REM === [1] Build log directory + paths ========================================
set "LOGDIR=%~dp0_Archive\scheduled_logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%" 2>nul

REM Rolling log = always-current, easy to find. Truncated each run.
set "ROLLLOG=%LOGDIR%\latest.log"
> "%ROLLLOG%" echo === Sale Report Scheduled Run ===
>> "%ROLLLOG%" echo Start: %DATE% %TIME%
>> "%ROLLLOG%" echo CWD:   %CD%
>> "%ROLLLOG%" echo User:  %USERNAME%
>> "%ROLLLOG%" echo Bat:   %~f0
>> "%ROLLLOG%" echo.

REM === [2] Compute timestamp via PowerShell (with stderr swallow + fallback) ==
set "TS="
for /f "usebackq tokens=*" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Date -Format 'yyyy-MM-dd_HHmmss'" 2^>nul`) do set "TS=%%i"
if "%TS%"=="" (
    >> "%ROLLLOG%" echo [WARN] PowerShell timestamp failed - using fallback timestamp
    set "TS=run"
)
>> "%ROLLLOG%" echo Timestamp: %TS%
>> "%ROLLLOG%" echo.

REM === [3] Find Python ========================================================
REM v3.10.1: Force UTF-8 for Python stdout/stderr (Python 3.x on Windows defaults to
REM cp1252 when stdout is redirected -> Unicode chars like '->' '!' crash the pipeline)
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

set PYTHON=
py --version >> "%ROLLLOG%" 2>&1
if %ERRORLEVEL% == 0 (
    set PYTHON=py
    goto :found
)
python --version >> "%ROLLLOG%" 2>&1
if %ERRORLEVEL% == 0 (
    set PYTHON=python
    goto :found
)
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
>> "%ROLLLOG%" echo [FATAL] Python not found in PATH or common install locations
goto :end_fail

:found
>> "%ROLLLOG%" echo Using Python: %PYTHON%
>> "%ROLLLOG%" echo.

REM === [4] Ensure deps (idempotent) ===========================================
%PYTHON% -c "import win32com.client" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    >> "%ROLLLOG%" echo Installing pywin32...
    %PYTHON% -m pip install pywin32 -q >> "%ROLLLOG%" 2>&1
)

%PYTHON% -c "import requests, dotenv" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    >> "%ROLLLOG%" echo Installing requests + python-dotenv...
    %PYTHON% -m pip install requests python-dotenv -q >> "%ROLLLOG%" 2>&1
)

REM === [5] Close any stale Excel ==============================================
taskkill /F /IM EXCEL.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

REM === [6] Run the pipeline ===================================================
>> "%ROLLLOG%" echo.
>> "%ROLLLOG%" echo === Starting daily_report.py --fetch-odoo ===
>> "%ROLLLOG%" echo.
%PYTHON% daily_report.py --fetch-odoo >> "%ROLLLOG%" 2>&1
set EXITCODE=%ERRORLEVEL%

>> "%ROLLLOG%" echo.
>> "%ROLLLOG%" echo === Finished at %DATE% %TIME% with exit code %EXITCODE% ===

REM === [7] Archive a timestamped copy
@echo off
setlocal
cd /d "%~dp0"
title YVF Adoption Dashboard - CS HAD

 echo ==============================================
 echo   YVF ADOPTION DASHBOARD - CS HAD
 echo ==============================================
 echo.

set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD (
    where python >nul 2>nul && set "PY_CMD=python"
)

if not defined PY_CMD (
    echo [ERROR] Python is not installed or not added to PATH.
    echo Please install Python 3.10 or later, then run this file again.
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating virtual environment...
    %PY_CMD% -m venv .venv
    if errorlevel 1 goto :error
)

echo [2/3] Checking required libraries...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo [3/3] Starting dashboard...
echo The dashboard will open automatically in your web browser.
echo Keep this window open while using the dashboard.
echo Press Ctrl+C in this window to stop it.
echo.
".venv\Scripts\python.exe" -m streamlit run app.py
exit /b 0

:error
echo.
echo [ERROR] The dashboard could not be started.
echo Please take a screenshot of this window and send it for checking.
echo.
pause
exit /b 1

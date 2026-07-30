@echo off
setlocal
cd /d "%~dp0"
title YVF Adoption Dashboard - CS HAD

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m streamlit run app.py
) else (
    echo The first-time setup has not been completed.
    echo Please double-click START_DASHBOARD.bat first.
    echo.
    pause
)

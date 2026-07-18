@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py
) else (
    py main.py
)

if errorlevel 1 (
    echo.
    echo TonBanAPP startup failed. Press any key to close this window.
    pause >nul
)

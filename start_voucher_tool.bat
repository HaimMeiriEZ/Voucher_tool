@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "voucher_app.py"
) else (
    python "voucher_app.py"
)

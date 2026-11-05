@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Sanal ortam olusturuluyor...
    py -3 -m venv .venv || python -m venv .venv
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
start "" ".venv\Scripts\pythonw.exe" run_panel.py

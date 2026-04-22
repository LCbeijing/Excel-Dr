@echo off
cd /d "%~dp0"
python -m pip install pyinstaller
python -m PyInstaller --noconsole --onefile cleaner.py --name Excel-Dr
pause

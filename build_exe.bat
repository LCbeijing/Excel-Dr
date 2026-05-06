@echo off
cd /d "%~dp0"
cargo build --manifest-path rust\excel_dr_core\Cargo.toml --release
if errorlevel 1 pause & exit /b 1
python -m pip install pyinstaller
python -m PyInstaller --noconsole --onefile cleaner.py --name Excel-Dr --clean
if errorlevel 1 pause & exit /b 1
copy /Y rust\excel_dr_core\target\release\excel_dr_core.exe dist\excel-dr-core.exe
pause

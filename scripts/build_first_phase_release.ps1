$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $root

cargo test --manifest-path ".\rust\excel_dr_core\Cargo.toml"
cargo build --manifest-path ".\rust\excel_dr_core\Cargo.toml" --release
python ".\scripts\smoke_test_current_backend.py"
python ".\scripts\smoke_test_rust_backend.py"
python -m PyInstaller --noconsole --onefile ".\cleaner.py" --name "Excel-Dr" --clean
Copy-Item -LiteralPath ".\rust\excel_dr_core\target\release\excel_dr_core.exe" -Destination ".\dist\excel-dr-core.exe" -Force

Get-ChildItem ".\dist" -Filter "*.exe" | Select-Object Name, Length, LastWriteTime

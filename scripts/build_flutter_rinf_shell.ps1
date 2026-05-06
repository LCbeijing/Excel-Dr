$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$app = Join-Path $root "apps\excel_dr_flutter"
$dist = Join-Path $root "dist"
$portableDir = Join-Path $dist "Excel-Dr-Flutter"
$zip = Join-Path $dist "Excel-Dr-Flutter-portable.zip"
$singleExe = Join-Path $dist "Excel-Dr-Single.exe"
$launcher = Join-Path $root "scripts\excel_dr_single_launcher.cs"
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

$env:Path = "C:\src\flutter\bin;C:\Program Files\CMake\bin;$env:Path"
$env:FLUTTER_STORAGE_BASE_URL = "https://storage.flutter-io.cn"
$env:PUB_HOSTED_URL = "https://pub.flutter-io.cn"

Set-Location $app

flutter pub get
rinf gen
cargo check --manifest-path ".\native\hub\Cargo.toml"
flutter build windows --release

$release = Resolve-Path ".\build\windows\x64\runner\Release"
New-Item -ItemType Directory -Path $dist -Force | Out-Null
$distFull = Resolve-Path $dist
$portableFull = [System.IO.Path]::GetFullPath($portableDir)
if (-not $portableFull.StartsWith($distFull.Path, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to write outside dist: $portableFull"
}
if (Test-Path -LiteralPath $portableFull) {
    Remove-Item -LiteralPath $portableFull -Recurse -Force
}
New-Item -ItemType Directory -Path $portableFull | Out-Null
Get-ChildItem -LiteralPath $release | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $portableFull -Recurse -Force
}
$innerExe = Join-Path $portableFull "excel_dr_flutter.exe"
if (Test-Path -LiteralPath $innerExe) {
    Rename-Item -LiteralPath $innerExe -NewName "Excel-Dr.exe"
}

if (Test-Path -LiteralPath $zip) {
    Remove-Item -LiteralPath $zip -Force
}
Compress-Archive -Path (Join-Path $portableFull "*") -DestinationPath $zip -Force

& $csc /nologo /target:winexe /platform:x64 /optimize+ /out:$singleExe /resource:$zip,ExcelDrPayload.zip /reference:System.Windows.Forms.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll $launcher

Get-ChildItem -LiteralPath $portableFull | Select-Object Name, Length, LastWriteTime
Get-FileHash -Algorithm SHA256 (Join-Path $portableFull "Excel-Dr.exe")
Get-FileHash -Algorithm SHA256 $zip
Get-FileHash -Algorithm SHA256 $singleExe

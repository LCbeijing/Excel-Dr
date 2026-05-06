$ErrorActionPreference = "Stop"

$env:Path = "C:\src\flutter\bin;C:\Program Files\CMake\bin;$env:Path"
$env:FLUTTER_STORAGE_BASE_URL = "https://storage.flutter-io.cn"
$env:PUB_HOSTED_URL = "https://pub.flutter-io.cn"

Write-Host "Flutter:" -ForegroundColor Cyan
flutter --version

Write-Host "`nRinf:" -ForegroundColor Cyan
rinf --help | Select-Object -First 8

Write-Host "`nDoctor:" -ForegroundColor Cyan
flutter doctor -v

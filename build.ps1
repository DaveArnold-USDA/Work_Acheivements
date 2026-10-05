param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

if ($Clean -and (Test-Path build)) { Remove-Item -Recurse -Force build }
if ($Clean -and (Test-Path dist)) { Remove-Item -Recurse -Force dist }

pyinstaller --noconfirm --clean --onefile --windowed --name WorkAchievementsLogger app.py
Write-Host "Built dist\WorkAchievementsLogger.exe"

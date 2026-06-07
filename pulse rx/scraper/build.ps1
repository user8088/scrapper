# Build a standalone Windows executable: dist\PulseRxScraper.exe
# Usage:  .\build.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "Installing build dependencies..."
python -m pip install --upgrade pip pyinstaller
python -m pip install -r requirements.txt

Write-Host "Building PulseRxScraper.exe..."
python -m PyInstaller --onefile --windowed --clean `
    --name PulseRxScraper `
    --hidden-import lxml `
    --hidden-import lxml._elementpath `
    app.py

Write-Host ""
Write-Host "Done -> $(Join-Path $PSScriptRoot 'dist\PulseRxScraper.exe')" -ForegroundColor Green

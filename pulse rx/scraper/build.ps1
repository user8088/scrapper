# Build a standalone Windows executable: dist\PulseRxScraper.exe
# Usage:  .\build.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "Installing build dependencies..."
python -m pip install --upgrade pip pyinstaller
python -m pip install -r requirements.txt

Write-Host "Building PulseRxScraper.exe..."
# PyInstaller logs progress to stderr; under Windows PowerShell that would trip a
# terminating error, so relax the preference and judge success by the exit code.
$ErrorActionPreference = "Continue"
python -m PyInstaller --onefile --windowed --clean `
    --name PulseRxScraper `
    --hidden-import lxml `
    --hidden-import lxml._elementpath `
    app.py
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

Write-Host ""
Write-Host "Done -> $(Join-Path $PSScriptRoot 'dist\PulseRxScraper.exe')" -ForegroundColor Green

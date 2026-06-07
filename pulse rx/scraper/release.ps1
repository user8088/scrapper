# Publish a new version. Bumps version.py, commits, tags, and pushes.
# GitHub Actions then builds the exe and publishes the release; every running
# install picks it up on next launch.
#
# Usage:  .\release.ps1 1.1.0
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version
)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$tag = "v$Version"

# Refuse if the tag already exists.
if (git tag --list $tag) {
    Write-Error "Tag $tag already exists. Pick a new version."
}

# Update version.py (the app's single source of truth).
Set-Content -Path (Join-Path $PSScriptRoot "version.py") `
    -Value "__version__ = `"$Version`"" -Encoding ascii

git add (Join-Path $PSScriptRoot "version.py")
git commit -m "Release v$Version"
git tag $tag

Write-Host "Pushing main and $tag..."
git push origin HEAD
git push origin $tag

Write-Host ""
Write-Host "Pushed $tag. GitHub Actions will build and publish the release." -ForegroundColor Green
Write-Host "Watch it at: https://github.com/user8088/scrapper/actions"

# Publish a new version. Bumps version.py, commits, tags, and pushes.
# GitHub Actions then builds the exe and publishes the release; every running
# install picks it up on next launch.
#
# Usage:  .\release.ps1 1.0.1
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version
)
Set-Location $PSScriptRoot

# git writes progress to stderr, which Windows PowerShell would treat as a
# terminating error under 'Stop'. Run native commands ourselves and judge them
# by their exit code instead.
$ErrorActionPreference = "Continue"
function Invoke-Git {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)
    & git @Args
    if ($LASTEXITCODE -ne 0) { throw "git $($Args -join ' ') failed (exit $LASTEXITCODE)" }
}

$tag = "v$Version"

# Refuse if the tag already exists.
if (git tag --list $tag) {
    throw "Tag $tag already exists. Pick a new version."
}

# Update version.py (the app's single source of truth).
Set-Content -Path (Join-Path $PSScriptRoot "version.py") `
    -Value "__version__ = `"$Version`"" -Encoding ascii

Invoke-Git add (Join-Path $PSScriptRoot "version.py")
Invoke-Git commit -m "Release v$Version"
Invoke-Git tag $tag

Write-Host "Pushing main and $tag..."
Invoke-Git push origin HEAD
Invoke-Git push origin $tag

Write-Host ""
Write-Host "Pushed $tag. GitHub Actions will build and publish the release." -ForegroundColor Green
Write-Host "Watch it at: https://github.com/user8088/scrapper/actions"

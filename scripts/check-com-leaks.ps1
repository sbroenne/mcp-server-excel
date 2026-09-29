#!/usr/bin/env pwsh
# Suspicious dynamic-object pattern guard, not proof of COM lifetime safety.

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "Scanning source for dynamic object access without a cleanup call..." -ForegroundColor Yellow

$leakFiles = @()
$cleanFiles = @()

$files = @(Get-ChildItem -LiteralPath (Join-Path $rootDir "src") -Recurse -File -Filter "*.cs" |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' -and $_.Name -notmatch '\.g\.cs$' })
if ($files.Count -eq 0) { throw 'No source files found for the dynamic cleanup pattern guard.' }
$files | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    $hasDynamic = $content -match "dynamic\s+\w+\s*=.*\."
    $hasRelease = $content -match "ComUtilities\.Release"
    $isSessionFile = $_.FullName -match "ExcelBatch\.cs|ExcelSession\.cs"

    $relativePath = $_.FullName.Replace("$rootDir\", "")

    if ($hasDynamic -and -not $hasRelease -and -not $isSessionFile) {
        $leakFiles += $_
        Write-Host "$relativePath - HAS COM objects but NO cleanup" -ForegroundColor Red
    } elseif ($hasDynamic -and $hasRelease) {
        $cleanFiles += $_
        Write-Host "$relativePath - Cleanup call present (not a lifetime proof)" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Clean files: $($cleanFiles.Count)" -ForegroundColor Green
Write-Host "  Leak files: $($leakFiles.Count)" -ForegroundColor Red

if ($leakFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "COM OBJECT LEAKS DETECTED!" -ForegroundColor Red
    Write-Host "Fix these files before committing:" -ForegroundColor Red
    $leakFiles | ForEach-Object {
        $rel = $_.FullName -replace [regex]::Escape("$rootDir\"), ''
        Write-Host "  - $rel" -ForegroundColor Red
    }
    exit 1
} else {
    Write-Host ""
    Write-Host "No suspicious dynamic cleanup patterns found; COM lifetime tests are still required." -ForegroundColor Green
    exit 0
}

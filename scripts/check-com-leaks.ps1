#!/usr/bin/env pwsh
# COM Object Leak Detection Script
# Run this before every commit to catch COM leaks

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "Scanning for COM object leaks..." -ForegroundColor Yellow

$leakFiles = @()
$cleanFiles = @()

Get-ChildItem -Path (Join-Path $rootDir "src") -Recurse -Filter "*.cs" | ForEach-Object {
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
        Write-Host "$relativePath - Proper COM cleanup" -ForegroundColor Green
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
    Write-Host "No COM object leaks detected!" -ForegroundColor Green
    exit 0
}

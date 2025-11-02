#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook to check for COM object leaks and Core Commands coverage

.DESCRIPTION
    Runs both COM leak checker and coverage audit before allowing commits.
    Ensures code quality and prevents coverage regression.

.EXAMPLE
    .\pre-commit.ps1

.NOTES
    This script is called by the Git pre-commit hook.
    To install: Copy .git/hooks/pre-commit (bash) or configure Git to use this PowerShell version.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "🔍 Checking for COM object leaks..." -ForegroundColor Cyan

try {
    $leakCheckScript = Join-Path $rootDir "scripts\check-com-leaks.ps1"
    & $leakCheckScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ COM object leaks detected! Fix them before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ COM leak check passed" -ForegroundColor Green
}
catch {
    Write-Host "⚠️  Error running COM leak check: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuing with coverage audit..." -ForegroundColor Gray
}

Write-Host ""
Write-Host "🔍 Checking Core Commands coverage..." -ForegroundColor Cyan

try {
    $auditScript = Join-Path $rootDir "scripts\audit-core-coverage.ps1"
    & $auditScript -FailOnGaps

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ Coverage gaps detected! All Core methods must be exposed via MCP Server." -ForegroundColor Red
        Write-Host "   Fix the gaps before committing (add enum values and mappings)." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ Coverage audit passed - 100% coverage maintained" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running coverage audit: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "✅ All pre-commit checks passed!" -ForegroundColor Green
exit 0

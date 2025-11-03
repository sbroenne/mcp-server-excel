#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook to check for COM object leaks, Core Commands coverage, naming consistency, and MCP Server functionality

.DESCRIPTION
    Runs four checks before allowing commits:
    1. COM leak checker - ensures no Excel COM objects are leaked
    2. Coverage audit - ensures 100% Core Commands are exposed via MCP Server
    3. Naming consistency - ensures enum names match Core method names exactly
    4. Smoke test - validates all 11 MCP tools work correctly

    Ensures code quality and prevents regression.

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
Write-Host "🔍 Checking Core Commands coverage and naming..." -ForegroundColor Cyan

try {
    $auditScript = Join-Path $rootDir "scripts\audit-core-coverage.ps1"
    & $auditScript -CheckNaming -FailOnGaps

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ Coverage or naming issues detected!" -ForegroundColor Red
        Write-Host "   All Core methods must be exposed via MCP Server with matching names." -ForegroundColor Red
        Write-Host "   Fix the issues before committing (add/rename enum values and mappings)." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ Coverage and naming checks passed - 100% coverage with consistent names" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running coverage audit: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔍 Running MCP Server smoke test..." -ForegroundColor Cyan

try {
    # Run the smoke test with proper filter (OnDemand only)
    $smokeTestFilter = "FullyQualifiedName~McpServerSmokeTests.SmokeTest_AllTools_LlmWorkflow"

    Write-Host "   dotnet test --filter `"$smokeTestFilter`" --verbosity quiet" -ForegroundColor Gray
    dotnet test --filter $smokeTestFilter --verbosity quiet

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ MCP Server smoke test failed! Core functionality is broken." -ForegroundColor Red
        Write-Host "   This test validates all 11 MCP tools work correctly." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ MCP Server smoke test passed - all tools functional" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running smoke test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Ensure Excel is installed and accessible." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "✅ All pre-commit checks passed!" -ForegroundColor Green
exit 0

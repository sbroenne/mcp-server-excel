#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Audit script to verify CLI action catalog stays in sync with Core action enums.

.DESCRIPTION
    Runs focused CLI tests that validate action mappings and the actions command output.
    Fails if no tests are executed (filter mismatch) or if any tests fail.

.EXAMPLE
    .\scripts\audit-cli-actions.ps1

.EXAMPLE
    .\scripts\audit-cli-actions.ps1 -Verbose
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "🔍 CLI Actions Audit" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan
Write-Host ""

$testProject = Join-Path $rootDir "tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj"
if (-not (Test-Path $testProject)) {
    Write-Host "❌ CLI test project not found: $testProject" -ForegroundColor Red
    exit 1
}

$filter = "FullyQualifiedName~ActionValidatorTests"
$verbosity = if ($VerbosePreference -ne "SilentlyContinue") { "normal" } else { "minimal" }

Write-Host "Running: dotnet test $testProject --filter \"$filter\"" -ForegroundColor Gray

$testOutput = dotnet test $testProject --filter $filter --verbosity $verbosity 2>&1 | Out-String
$testExitCode = $LASTEXITCODE

if (-not ($testOutput -match "Passed:\s*[1-9]")) {
    Write-Host "" 
    Write-Host "❌ CRITICAL: No CLI action tests passed! Filter may have matched zero tests." -ForegroundColor Red
    Write-Host "   Filter: $filter" -ForegroundColor Yellow
    Write-Host "   Verify tests exist: ExcelMcp.CLI.Tests.Unit.ActionValidatorTests" -ForegroundColor Yellow
    Write-Host "" 
    Write-Host $testOutput -ForegroundColor Gray
    exit 1
}

if ($testExitCode -ne 0) {
    Write-Host "" 
    Write-Host "❌ CLI actions audit failed." -ForegroundColor Red
    Write-Host "" 
    Write-Host $testOutput -ForegroundColor Gray
    exit 1
}

Write-Host "✅ CLI actions audit passed" -ForegroundColor Green
exit 0

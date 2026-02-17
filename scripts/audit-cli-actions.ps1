#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Audit script to verify CLI action catalog stays in sync with Core action enums.

.DESCRIPTION
    Delegates to check-cli-coverage.ps1 which auto-discovers all action enums
    and verifies each has a corresponding CLI command registered.

.EXAMPLE
    .\scripts\audit-cli-actions.ps1

.EXAMPLE
    .\scripts\audit-cli-actions.ps1 -Verbose
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "CLI Actions Audit" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan
Write-Host ""

# Delegate to the auto-discovering coverage script
$coverageScript = Join-Path $rootDir "scripts\check-cli-coverage.ps1"
& $coverageScript

exit $LASTEXITCODE

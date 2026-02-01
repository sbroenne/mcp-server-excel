#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Tests the Excel CLI end-to-end workflow - exactly what a user would do.

.DESCRIPTION
    This script demonstrates and tests a complete CLI workflow:
    1. Create session (auto-starts daemon, creates file)
    2. Set values
    3. Get values (verify roundtrip)
    4. Create a table
    5. List tables
    6. Create a chart
    7. List charts
    8. Close session (with save)
    9. Verify file exists

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [switch]$KeepFile  # Don't delete the test file after completion
)

$ErrorActionPreference = 'Stop'

# Find CLI executable (prefer Release build)
$cliPath = Join-Path $PSScriptRoot "..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"
if (-not (Test-Path $cliPath)) {
    $cliPath = Join-Path $PSScriptRoot "..\src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe"
}
if (-not (Test-Path $cliPath)) {
    Write-Error "CLI not found. Build first: dotnet build src/ExcelMcp.CLI"
    exit 1
}

$cli = (Resolve-Path $cliPath).Path
Write-Host "Using CLI: $cli" -ForegroundColor Cyan

# Generate unique test file
$testFile = Join-Path $env:TEMP "cli-workflow-test-$(Get-Random).xlsx"
Write-Host "Test file: $testFile" -ForegroundColor Cyan

$passed = 0
$failed = 0

function Test-Step {
    param(
        [string]$Name,
        [scriptblock]$Action,
        [scriptblock]$Verify = $null
    )

    Write-Host "`n[$Name]" -ForegroundColor Yellow
    try {
        $result = & $Action
        if ($Verify) {
            $verifyResult = & $Verify $result
            if (-not $verifyResult) {
                Write-Host "  FAIL: Verification failed" -ForegroundColor Red
                Write-Host "  Result: $result" -ForegroundColor Gray
                $script:failed++
                return $null
            }
        }
        Write-Host "  PASS" -ForegroundColor Green
        $script:passed++
        return $result
    }
    catch {
        Write-Host "  FAIL: $_" -ForegroundColor Red
        $script:failed++
        return $null
    }
}

# ============================================================================
# TEST WORKFLOW
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Excel CLI Workflow Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# 1. Create session (auto-starts daemon, creates file)
$session = Test-Step "Create session (create file)" {
    & $cli -q session create $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -ne $false
}

if (-not $session.sessionId) {
    Write-Host "`nFATAL: Could not open session. Aborting." -ForegroundColor Red
    exit 1
}

$sessionId = $session.sessionId
Write-Host "  Session ID: $sessionId" -ForegroundColor Gray

# 2. Set values
Test-Step "Set values (A1:C4)" {
    $values = '[["Product","Q1","Q2"],["Widget",100,150],["Gadget",80,90],["Device",200,180]]'
    & $cli -q range set-values --session $sessionId --sheet Sheet1 --range A1:C4 --values $values | ConvertFrom-Json
} -Verify {
    param($r)
    $r.ok -eq $true
}

# 3. Get values back
$getData = Test-Step "Get values (verify roundtrip)" {
    & $cli -q range get-values --session $sessionId --sheet Sheet1 --range A1:C4 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.ok -eq $true -and $r.d[0][0] -eq "Product"
}

Write-Host "  First cell: $($getData.d[0][0])" -ForegroundColor Gray

# 4. Create table (use --table not --name)
Test-Step "Create table 'SalesData'" {
    & $cli -q table create --session $sessionId --sheet Sheet1 --range A1:C4 --table SalesData --has-headers | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -or $r.tableName -eq "SalesData" -or $r.ok -eq $true
}

# 5. List tables
$tables = Test-Step "List tables" {
    & $cli -q table list --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.ok -eq $true -or $r.ts -ne $null
}

if ($tables.ts) {
    Write-Host "  Tables: $($tables.ts.Count)" -ForegroundColor Gray
}

# 6. Create chart
Test-Step "Create chart (below data)" {
    & $cli -q chart create-from-range --session $sessionId --sheet Sheet1 --source-range A1:C4 --chart-type column --chart "SalesChart" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.ok -eq $true -or $r.chartName -eq "SalesChart"
}

# 7. List charts
$charts = Test-Step "List charts" {
    & $cli -q chart list --session $sessionId --sheet Sheet1 | ConvertFrom-Json
} -Verify {
    param($r)
    if ($r -is [array]) { return $r.Count -ge 1 }
    return $null -ne $r
}

if ($charts) {
    $chartInfo = if ($charts -is [array]) { $charts[0] } else { $charts }
    Write-Host "  Chart: $($chartInfo.name) at $($chartInfo.topLeftCell)" -ForegroundColor Gray
}

# 8. Close session (with save)
Test-Step "Close session (with save)" {
    & $cli -q session close --session $sessionId --save | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 9. Verify file exists
Test-Step "Verify file exists" {
    if (Test-Path $testFile) {
        $size = (Get-Item $testFile).Length
        "File size: $size bytes"
    } else {
        throw "File not found"
    }
} -Verify {
    param($r)
    $r -match "bytes"
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed: $passed" -ForegroundColor Green
Write-Host "Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "Test file: $testFile" -ForegroundColor Gray

if (-not $KeepFile -and (Test-Path $testFile)) {
    Remove-Item $testFile -Force
    Write-Host "(Test file deleted)" -ForegroundColor Gray
} elseif ($KeepFile) {
    Write-Host "(Test file kept for inspection)" -ForegroundColor Yellow
}

if ($failed -gt 0) {
    Write-Host "`nSome tests FAILED!" -ForegroundColor Red
    exit 1
} else {
    Write-Host "`nAll tests PASSED!" -ForegroundColor Green
    exit 0
}

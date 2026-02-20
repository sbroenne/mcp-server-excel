#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Tests the Excel CLI end-to-end workflow - exactly what a user would do.

.DESCRIPTION
    This script demonstrates and tests a basic CLI workflow:
    1. Create session (auto-starts daemon, creates file)
    2. Create worksheet
    3. List worksheets
    4. Delete worksheet
    5. Close session (with save)
    6. Reopen saved file (session open - exercises Workbooks.Open path)
    7. List worksheets in reopened session
    8. Close reopened session
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

# 2. Create worksheet (simpler than set-values with JSON)
Test-Step "Create worksheet 'Data'" {
    & $cli -q sheet create --session $sessionId --sheet-name Data | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 3. List worksheets
$sheets = Test-Step "List worksheets" {
    & $cli -q sheet list --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -or $r.worksheets -ne $null
}

Write-Host "  Sheets: $(($sheets.worksheets | Measure-Object).Count)" -ForegroundColor Gray

# 4. Delete worksheet
Test-Step "Delete worksheet 'Data'" {
    & $cli -q sheet delete --session $sessionId --sheet-name Data | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 5. Close session (with save)
Test-Step "Close session (with save)" {
    & $cli -q session close --session $sessionId --save | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 6. Reopen saved file (session open - exercises Workbooks.Open path distinct from Add+SaveAs)
#    This step would catch deployment issues like missing office.dll (issue #487) because
#    ExcelBatch.ctor runs AutomationSecurity setup before opening any workbook.
$reopenSession = Test-Step "Reopen saved file (session open)" {
    & $cli -q session open $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -ne $false
}

# 6b. List worksheets in reopened session (proves the file loaded correctly)
if ($reopenSession -and $reopenSession.sessionId) {
    $reopenSessionId = $reopenSession.sessionId
    Test-Step "List worksheets in reopened session" {
        & $cli -q sheet list --session $reopenSessionId | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true -or $r.worksheets -ne $null
    }

    # 6c. Close reopened session
    Test-Step "Close reopened session" {
        & $cli -q session close --session $reopenSessionId | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true
    }
}

# 7. Verify file exists
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

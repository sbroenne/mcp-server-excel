#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Tests the Excel CLI end-to-end workflow - exactly what a user would do.

.DESCRIPTION
    This script demonstrates and tests a basic CLI workflow:
    1. Create session (auto-starts daemon, creates file)
    2. Create worksheet
    3. List worksheets
    4. Format multiple ranges
    5. Add and inspect a conditional-format rule with typed arguments
    6. Write a known value and create/delete a separate disposable worksheet
    7. Close session (with save)
    8. Reopen saved file (session open - exercises Workbooks.Open path)
    9. Read and verify the saved value in the reopened session
    10. Close reopened session
    11. Verify file exists

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1

.EXAMPLE
    .\scripts\Test-CliWorkflow.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [switch]$KeepFile,  # Don't delete the test file after completion
    [string]$PipeName
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

$previousPipeName = $env:EXCELMCP_CLI_PIPE
$selectedPipeName = if (-not [string]::IsNullOrWhiteSpace($PipeName)) {
    $PipeName
}
else {
    "excelmcp-cli-workflow-$PID-$([Guid]::NewGuid().ToString('N'))"
}
$env:EXCELMCP_CLI_PIPE = $selectedPipeName
Write-Host "Using private CLI pipe: $selectedPipeName" -ForegroundColor DarkGray

function Reset-CliWorkflowEnvironment {
    $cleanupExitCode = 0
    try {
        & (Join-Path $PSScriptRoot 'Stop-ExcelMcpProcesses.ps1') -PipeName $selectedPipeName
        $cleanupExitCode = $LASTEXITCODE
    }
    finally {
        if ($null -eq $previousPipeName) {
            Remove-Item Env:EXCELMCP_CLI_PIPE -ErrorAction SilentlyContinue
        }
        else {
            $env:EXCELMCP_CLI_PIPE = $previousPipeName
        }
    }

    return $cleanupExitCode
}

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
        $global:LASTEXITCODE = 0
        $result = & $Action
        if ($LASTEXITCODE -ne 0) { throw "Command failed with exit code $LASTEXITCODE. $($result | ConvertTo-Json -Depth 10)" }
        if ($result.success -eq $false -or $result.errorMessage) { throw "Command reported failure: $($result | ConvertTo-Json -Depth 10)" }
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

try {
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Excel CLI Workflow Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

:workflow do {
# 1. Create session (auto-starts daemon, creates file)
$session = Test-Step "Create session (create file)" {
    & $cli -q session create $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -eq $true
}

if (-not $session.sessionId) {
    Write-Host "`nFATAL: Could not open session. Aborting." -ForegroundColor Red
    $failed++
    break workflow
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
    $r.success -eq $true -and $r.worksheets -ne $null
}

Write-Host "  Sheets: $(($sheets.worksheets | Measure-Object).Count)" -ForegroundColor Gray

# 4. Format ranges (multi-value --range-addresses exercises string[] CLI option)
Test-Step "Format ranges on 'Data' (multi-value addresses)" {
    & $cli -q rangeformat format-ranges --session $sessionId --sheet-name Data --range-addresses "A1:A2" --range-addresses "C1:C2" --bold true --fill-color "#FFFF00" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 5. Add a conditional-format rule with typed integer/boolean arguments
Test-Step "Add typed conditional-format rule" {
    & $cli -q conditionalformat add-rule --session $sessionId --sheet-name Data --range-address "B1:B10" --rule-type top10 --rank 7 --top10-percent true --font-bold true --font-italic false | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

$conditionalFormatRules = Test-Step "Inspect typed conditional-format rule" {
    & $cli -q conditionalformat list-rules --session $sessionId --sheet-name Data --range-address "B1:B10" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.rules.Count -eq 1 -and
    $r.rules[0].top10.rank -eq 7 -and
    $r.rules[0].top10.percent -eq $true
}

# Keep Data for save/reopen verification; deletion uses a separate sheet.
Test-Step "Write persisted value" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address A1 --values '[[424242]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Protected write rejects occupied destination without mutation" {
    $rejected = & $cli -q range set-values --session $sessionId --sheet-name Data --range-address A1 --values '[[1]]' | ConvertFrom-Json
    if ($LASTEXITCODE -ne 1 -or $rejected.success -ne $false -or $rejected.errorCategory -ne 'Conflict' -or $rejected.errorMessage -notmatch '\$A\$1') {
        throw "Expected a categorized occupied-cell failure: $($rejected | ConvertTo-Json -Depth 10)"
    }
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address A1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][0] -eq 424242
}
Test-Step "Explicit allow permits intentional replacement" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address A1 --values '[[424242]]' --overwrite-policy allow | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Create disposable worksheet" {
    & $cli -q sheet create --session $sessionId --sheet-name Disposable | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Delete disposable worksheet" {
    & $cli -q sheet delete --session $sessionId --sheet-name Disposable | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 7. Close session (with save)
Test-Step "Close session (with save)" {
    & $cli -q session close --session $sessionId --save | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 8. Reopen saved file (session open - exercises Workbooks.Open path distinct from Add+SaveAs)
#    This step would catch deployment issues like missing office.dll (issue #487) because
#    ExcelBatch.ctor runs AutomationSecurity setup before opening any workbook.
$reopenSession = Test-Step "Reopen saved file (session open)" {
    & $cli -q session open $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -eq $true
}

# 9. List worksheets in reopened session (proves the file loaded correctly)
if ($reopenSession -and $reopenSession.sessionId) {
    $reopenSessionId = $reopenSession.sessionId
    Test-Step "Verify value survived save and reopen" {
        & $cli -q range get-values --session $reopenSessionId --sheet-name Data --range-address A1 | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true -and $r.values.Count -eq 1 -and $r.values[0][0] -eq 424242
    }

    # 10. Close reopened session
    Test-Step "Close reopened session" {
        & $cli -q session close --session $reopenSessionId | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true
    }
}

# 11. Verify file exists
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
} while ($false)

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed: $passed" -ForegroundColor Green
Write-Host "Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "Test file: $testFile" -ForegroundColor Gray

if ($KeepFile) {
    Write-Host "(Test file kept for inspection)" -ForegroundColor Yellow
}

if ($failed -gt 0) {
    Write-Host "`nSome tests FAILED!" -ForegroundColor Red
    $workflowExitCode = 1
} else {
    Write-Host "`nAll tests PASSED!" -ForegroundColor Green
    $workflowExitCode = 0
}
}
finally {
    try { $cleanupExitCode = Reset-CliWorkflowEnvironment }
    finally {
        if (-not $KeepFile -and (Test-Path -LiteralPath $testFile)) {
            Remove-Item -LiteralPath $testFile -Force
        }
    }
}

if ($cleanupExitCode -ne 0) {
    Write-Error "Owned CLI cleanup failed with exit code $cleanupExitCode."
    exit $cleanupExitCode
}

exit $workflowExitCode

#Requires -Version 7.0
<#
.SYNOPSIS
Tests ordinary temporary XLSX fixtures opened through macOS LaunchServices.
.DESCRIPTION
Never accesses Excel's container or requests permissions. Requires existing
Automation consent and running Excel. Copies an intact Excel-authored workbook
template; all editing, calculation and
persistence assertions use real Excel.
Does not establish Excel-native workbook creation or production MCP/CLI parity.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) { throw 'The file handoff spike requires macOS.' }
. (Join-Path $PSScriptRoot 'MacTestEnvironment.ps1')
Assert-MacAutomationAllowed
$suite = [Diagnostics.Stopwatch]::StartNew()
$runId = [guid]::NewGuid().ToString('N')
$directory = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-handoff-$runId"
$owned = @((Join-Path $directory 'main.xlsx'), (Join-Path $directory 'sentinel.xlsx'))
$bridge = Join-Path $PSScriptRoot 'ExcelSpike.applescript'
$completed = $false
$timedOut = $false

function Invoke-HandoffProcess {
    param([string]$Executable, [string[]]$Arguments, [int]$ExpectedErrorCode = 0)
    $remaining = [Math]::Min(45000, 150000 - $suite.ElapsedMilliseconds)
    if ($remaining -le 0) {
        $script:timedOut = $true
        throw 'File-handoff suite exceeded its 150-second deadline.'
    }
    $info = [Diagnostics.ProcessStartInfo]::new($Executable)
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $info.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw 'Could not start the handoff operation.' }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit([int]$remaining)) {
            $script:timedOut = $true
            $process.Kill()
            $process.WaitForExit()
            throw 'Handoff operation timed out. Excel was not stopped; no further automation will run.'
        }
        $outputText = $stdout.GetAwaiter().GetResult()
        $errorText = $stderr.GetAwaiter().GetResult()
        if ($errorText -match '\(-1712\)') { $script:timedOut = $true }
        if ($ExpectedErrorCode -ne 0) {
            if ($process.ExitCode -eq 0 -or $errorText -notmatch "\($ExpectedErrorCode\)") {
                throw "Expected error $ExpectedErrorCode, received: $errorText"
            }
            return @{ errorCode = $ExpectedErrorCode } | ConvertTo-Json -Compress
        }
        if ($process.ExitCode -ne 0) { throw "Handoff operation failed: $errorText" }
        return $outputText
    }
    finally { $process.Dispose() }
}

function Invoke-HandoffProbe {
    param([string]$Action, [string]$Path, [int]$ExpectedErrorCode = 0)
    $result = Invoke-HandoffProcess /usr/bin/osascript @($bridge, $Action, $Path) $ExpectedErrorCode |
        ConvertFrom-Json -AsHashtable
    if ($result.ContainsKey('success') -and (-not $result.success -or $result.errorMessage -ne '')) {
        throw "Invalid success/error result from $Action."
    }
    return $result
}

function Open-HandoffFixture {
    param([string]$Path)
    $null = Invoke-HandoffProcess /usr/bin/open @('-g', '-b', 'com.microsoft.Excel', $Path)
    $null = Invoke-HandoffProbe 'wait-open' $Path
}

function New-BlankFixture {
    param([string]$Path)
    $template = Join-Path $PSScriptRoot '../../../src/ExcelMcp.Service/Mac/Templates/Blank.xlsx'
    Copy-Item -LiteralPath $template -Destination $Path
}

function Assert-HandoffData {
    param([hashtable]$Data)
    if ($Data.values.Count -ne 3 -or @($Data.values | Where-Object Count -NE 2).Count -ne 0 -or
        $Data.values[0][0] -cne 'Item' -or $Data.values[0][1] -cne 'Amount' -or
        $Data.values[1][0] -cne 'Alpha' -or $Data.values[1][1] -ne 10 -or
        $Data.values[2][0] -cne 'Beta' -or $Data.values[2][1] -ne 20 -or
        $Data.calculated.Count -ne 1 -or $Data.calculated[0].Count -ne 1 -or $Data.calculated[0][0] -ne 30 -or
        $Data.formulas.Count -ne 1 -or $Data.formulas[0].Count -ne 1 -or $Data.formulas[0][0] -cne '=SUM(B2:B3)' -or
        $Data.numberFormat -ne '0.00' -or $Data.text -cne "quote `" slash \ newline`n$([char]937)") {
        throw 'Real Excel data, shapes, formulas, formatting or text did not match.'
    }
}

[void][IO.Directory]::CreateDirectory($directory)
try {
    foreach ($path in $owned) {
        New-BlankFixture $path
        Open-HandoffFixture $path
        $null = Invoke-HandoffProbe 'prepare-existing' $path
        Open-HandoffFixture $path
    }
    Assert-HandoffData (Invoke-HandoffProbe 'read' $owned[0])
    $null = Invoke-HandoffProbe 'missing-sheet' $owned[0] 9006
    Assert-HandoffData (Invoke-HandoffProbe 'read' $owned[0])
    $bulk = Invoke-HandoffProbe 'bulk' $owned[0]
    if ($bulk.values.Count -ne 1000 -or $bulk.total -ne 5005000) { throw 'Bulk rows/calculation did not match.' }
    for ($row = 0; $row -lt 1000; $row++) {
        if ($bulk.values[$row].Count -ne 10) { throw 'Bulk columns did not match.' }
        for ($column = 0; $column -lt 10; $column++) {
            if ($bulk.values[$row][$column] -ne ($row + 1) * ($column + 1)) { throw 'Bulk value did not match.' }
        }
    }
    $null = Invoke-HandoffProbe 'discard' $owned[0]
    Assert-HandoffData (Invoke-HandoffProbe 'read' $owned[1])
    Open-HandoffFixture $owned[0]
    Assert-HandoffData (Invoke-HandoffProbe 'read' $owned[0])
    foreach ($path in $owned) { $null = Invoke-HandoffProbe 'cleanup' $path }
    foreach ($path in $owned) {
        if (([IO.File]::GetAttributes($path) -band [IO.FileAttributes]::Directory) -ne 0) {
            throw 'A fixture file was replaced by a directory; refusing cleanup.'
        }
        Remove-Item -LiteralPath $path
    }
    [IO.Directory]::Delete($directory)
    $completed = $true
    [ordered]@{
        success = $true
        errorMessage = ''
        fixtureLocation = 'ordinary temporary directory'
        openingMechanism = 'LaunchServices'
        bulkCellsVerified = 10000
        elapsedMs = $suite.ElapsedMilliseconds
        checks = @('open-existing', 'real-Excel-edit-save', 'explicit-workbook-targeting',
            '2d-values', '1x1-formula-result', 'text-round-trip', 'number-format',
            'missing-sheet-error', 'read-after-error',
            'bulk-values-and-calculation', 'discard', 'save-reopen', 'sentinel-preserved', 'owned-cleanup')
        scope = 'Opaque copies of an Excel-authored template; not production MCP/CLI E2E.'
    } | ConvertTo-Json -Depth 5
}
finally {
    if (-not $completed) {
        Write-Warning "Incomplete handoff probe; fixtures retained at $directory. Timeout=$timedOut. No automatic retry or further Excel cleanup was attempted."
    }
}

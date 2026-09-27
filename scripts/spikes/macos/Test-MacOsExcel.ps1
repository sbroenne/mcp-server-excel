#Requires -Version 7.0
<#
.SYNOPSIS
Runs the ordinary-file LaunchServices spike by default.
.DESCRIPTION
Default mode uses temporary OOXML fixtures without accessing Excel's container.
The explicit AllowExcelContainerAccess switch reproduces the older container
experiment, which is NOT prompt-free. That mode creates two disposable workbooks inside Excel's sandbox, checks native automation
across separate osascript processes, and removes only those workbooks. Does not quit Excel,
change macro security, or change application-wide calculation/alert settings.
This is not an MCP/CLI implementation or a substitute for Windows COM E2E.
#>
[CmdletBinding()]
param(
    [switch]$AllowExcelContainerAccess
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) {
    throw 'This spike requires macOS, PowerShell 7, and installed desktop Excel.'
}
if (-not $AllowExcelContainerAccess) {
    & (Join-Path $PSScriptRoot 'Test-MacFileHandoff.ps1')
    return
}
. (Join-Path $PSScriptRoot 'MacTestEnvironment.ps1')
Assert-MacAutomationAllowed

$scriptPath = Join-Path $PSScriptRoot 'ExcelSpike.applescript'
$results = [System.Collections.Generic.List[object]]::new()
$runId = [guid]::NewGuid().ToString('N')
$excelDocuments = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Library/Containers/com.microsoft.Excel/Data/Documents'
if (-not (Test-Path -LiteralPath $excelDocuments -PathType Container)) {
    throw 'Excel sandbox Documents folder was not found. Launch desktop Excel before running the spike.'
}
$tempDirectory = Join-Path $excelDocuments "excelmcp-macos-spike-$runId"
$mainPath = Join-Path $tempDirectory "main workbook $runId.xlsx"
$sentinelPath = Join-Path $tempDirectory "sentinel workbook $runId.xlsx"
$timedOut = $false

function Invoke-ExcelProbe {
    param(
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Path,
        [switch]$ExpectFailure
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new('/usr/bin/osascript')
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.ArgumentList.Add($scriptPath)
    $startInfo.ArgumentList.Add($Action)
    $startInfo.ArgumentList.Add($Path)
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(45000)) {
            $script:timedOut = $true
            $process.Kill()
            $process.WaitForExit()
            throw 'Apple Events host timed out. Excel may still be busy; no further automation will be attempted.'
        }
        $outputText = $stdout.GetAwaiter().GetResult()
        $errorText = $stderr.GetAwaiter().GetResult()
        if ($errorText -match '\(-1712\)') {
            $script:timedOut = $true
            throw 'Excel reported an Apple Events timeout (-1712); no further automation will be attempted.'
        }
        if ($ExpectFailure) {
            if ($process.ExitCode -eq 0 -or [string]::IsNullOrWhiteSpace($errorText)) {
                throw "Expected a nonzero exit and error text for $Action."
            }
            # An unrelated compile/permission error must not pass the negative test.
            if ($errorText -notmatch '\(9006\)') {
                throw "Expected explicit missing-worksheet error 9006, received: $errorText"
            }
            $data = @{ errorCode = 9006 }
        }
        else {
            if ($process.ExitCode -ne 0) {
                throw "AppleScript action '$Action' failed: $errorText"
            }
            $data = ConvertFrom-Json -InputObject $outputText -AsHashtable
            if ($data.ContainsKey('success') -and (-not $data.success -or $data.errorMessage -ne '')) {
                throw "Invalid success/error result for $Action."
            }
        }
        $timer.Stop()
        Write-Verbose "$Action completed in $($timer.ElapsedMilliseconds) ms."
        $results.Add([ordered]@{ action = $Action; elapsedMs = $timer.ElapsedMilliseconds })
        return $data
    }
    finally {
        $process.Dispose()
    }
}

function Assert-WorkbookData {
    param([Parameter(Mandatory)][hashtable]$Data)

    $actual = ConvertTo-Json -InputObject $Data.values -Compress -Depth 5
    if ($actual -ne '[["Item","Amount"],["Alpha",10.0],["Beta",20.0]]' -and
        $actual -ne '[["Item","Amount"],["Alpha",10],["Beta",20]]') {
        throw "Unexpected range matrix: $actual"
    }
    if ($Data.calculated.Count -ne 1 -or $Data.calculated[0].Count -ne 1 -or $Data.calculated[0][0] -ne 30) {
        throw 'Single-cell result must be a 1x1 matrix containing 30.'
    }
    if ($Data.formulas.Count -ne 1 -or $Data.formulas[0].Count -ne 1 -or $Data.formulas[0][0] -ne '=SUM(B2:B3)') {
        throw 'Formula round-trip failed.'
    }
    $expectedText = "quote `" slash \ newline`n$([char]937)"
    if ($Data.text -cne $expectedText -or $Data.numberFormat -ne '0.00') {
        throw 'Text escaping/Unicode or number-format round-trip failed.'
    }
}

[void][System.IO.Directory]::CreateDirectory($tempDirectory)
$completed = $false
$primaryFailure = $null
try {
    $version = Invoke-ExcelProbe -Action version -Path $mainPath
    $null = Invoke-ExcelProbe -Action create -Path $mainPath
    $null = Invoke-ExcelProbe -Action create -Path $sentinelPath
    Assert-WorkbookData (Invoke-ExcelProbe -Action read -Path $mainPath)
    $null = Invoke-ExcelProbe -Action missing-sheet -Path $mainPath -ExpectFailure
    Assert-WorkbookData (Invoke-ExcelProbe -Action read -Path $mainPath)

    $bulk = Invoke-ExcelProbe -Action bulk -Path $mainPath
    if ($bulk.values.Count -ne 1000 -or $bulk.total -ne 5005000) {
        throw "Bulk mismatch: rows=$($bulk.values.Count), total=$($bulk.total); expected 1000 rows and 5005000."
    }
    for ($row = 0; $row -lt 1000; $row++) {
        if ($bulk.values[$row].Count -ne 10) { throw "Bulk column count differs in row $row." }
        for ($column = 0; $column -lt 10; $column++) {
            if ($bulk.values[$row][$column] -ne ($row + 1) * ($column + 1)) {
                throw "Bulk cell mismatch at row $row column $column."
            }
        }
    }
    $null = Invoke-ExcelProbe -Action discard -Path $mainPath
    Assert-WorkbookData (Invoke-ExcelProbe -Action read -Path $sentinelPath)
    if (-not (Test-Path -LiteralPath $mainPath -PathType Leaf)) { throw 'Excel did not save the workbook.' }
    $null = Invoke-ExcelProbe -Action open -Path $mainPath
    Assert-WorkbookData (Invoke-ExcelProbe -Action read -Path $mainPath)
    $null = Invoke-ExcelProbe -Action close -Path $mainPath
    Assert-WorkbookData (Invoke-ExcelProbe -Action read -Path $sentinelPath)
    $completed = $true
}
catch {
    $primaryFailure = $_.Exception.Message
    Write-Error -Message $primaryFailure -ErrorAction Continue
    throw
}
finally {
    if ($timedOut) {
        Write-Warning "Timeout: retained disposable workbooks at $tempDirectory for manual recovery; Excel was not killed."
    }
    else {
        $cleanupErrors = [System.Collections.Generic.List[string]]::new()
        foreach ($path in @($mainPath, $sentinelPath)) {
            try {
                $null = Invoke-ExcelProbe -Action cleanup -Path $path
            }
            catch {
                $cleanupErrors.Add($_.Exception.Message)
            }
            if ($timedOut) { break }
        }
        if ($cleanupErrors.Count -gt 0) {
            throw "Initial failure: $primaryFailure. Cleanup failed; retained $tempDirectory. $($cleanupErrors -join '; ')"
        }
        if ($completed) {
            foreach ($path in @($mainPath, $sentinelPath)) {
                if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
            }
            # Non-recursive: unexpected files are retained rather than deleted.
            try {
                [System.IO.Directory]::Delete($tempDirectory)
            }
            catch {
                throw "Could not remove disposable directory: $($_.Exception.Message)"
            }
        }
        else {
            Write-Warning "Failed run retained at $tempDirectory for investigation."
        }
    }
}

if ($completed) {
    [ordered]@{
        success = $true
        errorMessage = ''
        excelVersion = $version.version
        architecture = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
        bulkCellsVerified = 10000
        checks = @(
            'create-save', 'cross-process-read', 'explicit-workbook-targeting',
            '2d-values', '1x1-formula-result', 'text-round-trip', 'number-format',
            'missing-sheet-error', 'read-after-error', 'bulk-values-and-calculation',
            'discard-unsaved-changes', 'save-reopen', 'sentinel-preserved', 'owned-cleanup'
        )
        calls = $results.ToArray()
    } | ConvertTo-Json -Depth 6
}
$global:LASTEXITCODE = 0

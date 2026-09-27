#Requires -Version 7.0
<#
.SYNOPSIS
Compares .NET, Swift, and the current JXA host without requesting consent.
.DESCRIPTION
Does not send Excel commands, launch Excel, or access Excel's container.
The compiled Swift probe is a disposable, unsigned spike, not a production helper.
Equal results do not establish permission identity for MCP, CLI, or osascript.
.PARAMETER ManagedProbe
Internal child-process mode, also usable as a bounded standalone .NET preflight.
#>
[CmdletBinding()]
param([switch]$ManagedProbe)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) { throw 'The automation boundary spike requires macOS.' }

if ($ManagedProbe) {
    Add-Type -Path (Join-Path $PSScriptRoot '../../../src/ExcelMcp.Service/Mac/MacAutomationAccess.cs')
    $status = [Sbroenne.ExcelMcp.Service.Mac.MacAutomationAccess]::Check()
    [ordered]@{
        implementation = '.NET'
        status = [Sbroenne.ExcelMcp.Service.Mac.MacAutomationAccess]::DescribeStatus($status)
        osStatus = $status
        descriptorSize = [Runtime.InteropServices.Marshal]::SizeOf(
            [type][Sbroenne.ExcelMcp.Service.Mac.MacAutomationAccess+Descriptor])
        requestedConsent = $false
    } | ConvertTo-Json -Compress
    if ($status -ne 0) { exit 2 }
    exit 0
}

function Invoke-BoundedProcess {
    param([string]$Executable, [string[]]$Arguments, [int]$TimeoutSeconds)
    $info = [Diagnostics.ProcessStartInfo]::new($Executable)
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $info.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        if (-not $process.Start()) { throw "Could not start $Executable." }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "Boundary probe timed out after $TimeoutSeconds seconds; Excel was not stopped."
        }
        return @{
            exitCode = $process.ExitCode
            stdout = $stdout.GetAwaiter().GetResult()
            stderr = $stderr.GetAwaiter().GetResult()
            elapsedMs = $timer.ElapsedMilliseconds
        }
    }
    finally { $process.Dispose() }
}

$directory = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-boundary-$([guid]::NewGuid().ToString('N'))"
$binary = Join-Path $directory 'permission-probe'
[void][IO.Directory]::CreateDirectory($directory)
try {
    $build = Invoke-BoundedProcess /usr/bin/xcrun @(
        'swiftc', (Join-Path $PSScriptRoot 'MacAutomationProbe.swift'), '-o', $binary
    ) 120
    if ($build.exitCode -ne 0) { throw "Swift probe build failed: $($build.stderr)" }
    $managed = Invoke-BoundedProcess (Join-Path $PSHOME 'pwsh') @(
        '-NoProfile', '-File', $PSCommandPath, '-ManagedProbe'
    ) 15
    $native = Invoke-BoundedProcess $binary @() 15
    $jxa = Invoke-BoundedProcess /usr/bin/osascript @(
        '-l', 'JavaScript', (Join-Path $PSScriptRoot 'MacAutomationProbe.js')
    ) 15
    $results = foreach ($result in @($managed, $native, $jxa)) {
        if ($result.exitCode -notin @(0, 2)) { throw "Permission probe failed: $($result.stderr)" }
        $parsed = ConvertFrom-Json $result.stdout
        if ($parsed.requestedConsent -ne $false) { throw 'Probe violated the no-consent contract.' }
        $parsed | Add-Member -NotePropertyName elapsedMs -NotePropertyValue $result.elapsedMs -PassThru
    }
    if ($results[0].descriptorSize -ne $results[1].descriptorSize) {
        throw 'Managed descriptor size does not match the native SDK.'
    }
    [ordered]@{
        permissionChecksAgree = @($results.status | Select-Object -Unique).Count -eq 1
        readyForAutomation = @($results | Where-Object status -NE 'Allowed').Count -eq 0
        checks = @($results)
        scope = 'Permission checks only; no Excel commands or protected filesystem access.'
    } | ConvertTo-Json -Depth 5
    if (@($results | Where-Object status -NE 'Allowed').Count -ne 0) { exit 2 }
}
finally {
    if (Test-Path -LiteralPath $binary) { Remove-Item -LiteralPath $binary }
    [IO.Directory]::Delete($directory)
}

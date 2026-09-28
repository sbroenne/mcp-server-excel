#Requires -Version 7.0
<#
.SYNOPSIS
Tests the experimental macOS workbook slice through the real CLI and MCP processes.
.DESCRIPTION
Requires running desktop Excel and existing Automation consent. Uses ordinary
temporary fixtures and LaunchServices; never requests permission or accesses
Excel's container. Includes repository-authored MS-QDEFF/OOXML Power Query
fixtures. Refresh and VBA remain gated unless their explicit assertions pass.
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [string]$PipeName,
    [switch]$IncludePowerQueryFixtures
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) { throw 'This runner requires macOS desktop Excel.' }
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'spikes/macos/MacTestEnvironment.ps1')
Assert-MacAutomationAllowed

$pipe = if ([string]::IsNullOrWhiteSpace($PipeName)) { "em-$([guid]::NewGuid().ToString('N'))" } else { $PipeName }
if ([Text.Encoding]::UTF8.GetByteCount((Join-Path ([IO.Path]::GetTempPath()) "CoreFxPipe_$pipe")) -gt 104) {
    throw 'The test pipe exceeds the macOS domain-socket path limit. Supply a shorter PipeName.'
}

function Invoke-MacTestCommand {
    param([string]$Executable, [string[]]$Arguments, [int]$TimeoutSeconds, [hashtable]$Environment = @{})
    $start = [Diagnostics.ProcessStartInfo]::new($Executable)
    $start.WorkingDirectory = $root
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $start.ArgumentList.Add($argument) }
    foreach ($entry in $Environment.GetEnumerator()) { $start.Environment[$entry.Key] = $entry.Value }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw "Could not start $Executable." }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "Test command exceeded $TimeoutSeconds seconds. Shared Excel was not stopped."
        }
        return @{
            exitCode = $process.ExitCode
            stdout = $stdout.GetAwaiter().GetResult()
            stderr = $stderr.GetAwaiter().GetResult()
        }
    }
    finally { $process.Dispose() }
}

if (-not $SkipBuild) {
    $build = Invoke-MacTestCommand dotnet @(
        'build', 'Sbroenne.ExcelMcp.sln', '-c', 'Release',
        '-p:EnableWindowsTargeting=true', '--nologo', '-v', 'minimal'
    ) 300
    Write-Host $build.stdout
    if ($build.exitCode -ne 0) { throw "Release build failed: $($build.stderr)" }
}

$cli = Join-Path $root 'src/ExcelMcp.CLI/bin/Release/net10.0/excelcli'
$runtimes = Invoke-MacTestCommand dotnet @('--list-runtimes') 30
$runtimePaths = [regex]::Matches($runtimes.stdout, '(?m)^Microsoft\.NETCore\.App \S+ \[(.+)\]\r?$')
if ($runtimes.exitCode -ne 0 -or $runtimePaths.Count -eq 0) {
    throw "Cannot locate the .NET runtime for the CLI apphost: $($runtimes.stderr)"
}
# Homebrew's executable directory is not necessarily its runtime root.
$dotnetRoot = Split-Path -Parent (Split-Path -Parent $runtimePaths[0].Groups[1].Value)
$environment = @{
    EXCELMCP_MAC_E2E = '1'
    EXCELMCP_MAC_E2E_PIPE = $pipe
    EXCELMCP_CLI_PIPE = $pipe
    DOTNET_ROOT = $dotnetRoot
}
if ($IncludePowerQueryFixtures) {
    $environment.EXCELMCP_MAC_PQ_FIXTURE_E2E = '1'
}
try {
    $test = Invoke-MacTestCommand dotnet @(
        'test', 'tests/ExcelMcp.Portable.Tests/ExcelMcp.Portable.Tests.csproj',
        '-c', 'Release', '--no-build', '--nologo', '-v', 'minimal',
        '--filter', 'FullyQualifiedName~MacExcelE2ETests', '--blame-hang-timeout', '5m'
    ) 600 $environment
    Write-Host $test.stdout
    if (-not [string]::IsNullOrWhiteSpace($test.stderr)) { Write-Host $test.stderr }
    $expectedPassed = if ($IncludePowerQueryFixtures) { 6 } else { 4 }
    $expectedSkipped = if ($IncludePowerQueryFixtures) { 0 } else { 1 }
    $expectedTotal = if ($IncludePowerQueryFixtures) { 6 } else { 5 }
    $summaryPattern = "Passed!.*Failed:\s*0\b.*Passed:\s*$expectedPassed\b.*Skipped:\s*$expectedSkipped\b.*Total:\s*$expectedTotal\b"
    if ($test.exitCode -ne 0 -or $test.stdout -notmatch $summaryPattern) {
        throw "Expected $expectedPassed passed and $expectedSkipped skipped macOS workflows; missing or failed cases are not success."
    }
}
finally {
    if ([IO.File]::Exists($cli)) {
        $cleanup = Invoke-MacTestCommand $cli @('-q', 'service', 'stop') 30 $environment
        if ($cleanup.exitCode -ne 0 -or -not ($cleanup.stdout | ConvertFrom-Json).success) {
            throw "Private test daemon cleanup failed: $($cleanup.stdout) $($cleanup.stderr)"
        }
    }
}
if ($IncludePowerQueryFixtures) {
    Write-Host 'macOS CLI/MCP workbook and repository-owned Power Query fixture slices passed. Refresh, Windows COM, and unsupported Mac features remain separate gates.'
} else {
    Write-Host 'macOS CLI/MCP workbook slice passed. Power Query fixture acceptance requires -IncludePowerQueryFixtures and remains gated.'
}
$global:LASTEXITCODE = 0

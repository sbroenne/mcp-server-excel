#Requires -Version 7.0
<#
.SYNOPSIS
Launches the native development helper through LaunchServices, not as a shell child.
.DESCRIPTION
Check never requests consent or touches Excel's container. Setup requires
explicit permission to prompt. Run uses only the helper's existing setup.
The app bounds its work to 120 seconds; the launcher has a 150-second deadline.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ApplicationPath,
    [ValidateSet('check', 'setup', 'run')][string]$Mode = 'check',
    [switch]$AllowPermissionPrompts
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) { throw 'The native helper requires macOS.' }
if ($Mode -eq 'setup' -and -not $AllowPermissionPrompts) {
    throw 'Native helper setup requires explicit -AllowPermissionPrompts consent.'
}
if ($Mode -ne 'setup' -and $AllowPermissionPrompts) {
    throw 'Only explicit setup can request permission prompts.'
}
$bundle = [IO.Path]::GetFullPath($ApplicationPath)
if (-not [IO.File]::Exists((Join-Path $bundle 'Contents/MacOS/ExcelMcpMacPermissionProbe'))) {
    throw 'Build the native probe app first; the expected executable is missing.'
}
$resultPath = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-native-result-$([guid]::NewGuid().ToString('N')).json"
$info = [Diagnostics.ProcessStartInfo]::new('/usr/bin/open')
$info.UseShellExecute = $false
$info.RedirectStandardOutput = $true
$info.RedirectStandardError = $true
foreach ($argument in @('-n', '-W', $bundle, '--args', '--mode', $Mode, '--output', $resultPath)) {
    $info.ArgumentList.Add($argument)
}
$process = [Diagnostics.Process]::new()
$process.StartInfo = $info
try {
    if (-not $process.Start()) { throw 'Could not start the native helper launcher.' }
    $stdout = $process.StandardOutput.ReadToEndAsync()
    $stderr = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(150000)) {
        $process.Kill()
        $process.WaitForExit()
        throw 'Helper launcher timed out; only the owned launcher was stopped. The helper has its own 120-second deadline; Excel was not stopped.'
    }
    $outputText = $stdout.GetAwaiter().GetResult()
    $errorText = $stderr.GetAwaiter().GetResult()
    if ($process.ExitCode -ne 0) { throw "LaunchServices failed: $errorText $outputText" }
    if (-not [IO.File]::Exists($resultPath)) {
        throw 'Native helper exited without a result. Inspect macOS diagnostics; do not infer success.'
    }
    $report = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
    $report | ConvertTo-Json -Depth 8
    if (-not $report.success) { exit 2 }
}
finally {
    $process.Dispose()
    if ([IO.File]::Exists($resultPath)) { Remove-Item -LiteralPath $resultPath }
}

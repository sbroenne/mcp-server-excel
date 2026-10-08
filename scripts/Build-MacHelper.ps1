#Requires -Version 7.0
<#
.SYNOPSIS
Prepares independently versioned helper source or builds the .xlam using desktop Excel.
.DESCRIPTION
PrepareOnly emits the reviewed import module without opening Excel. The first
bootstrap import is manual. Building requires that Excel-authored bootstrap to
already be open with its macros approved. It does not change VBA trust settings.
#>
[CmdletBinding()]
param(
    [switch]$PrepareOnly,
    [string]$BootstrapPath,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../artifacts/mac-helper'),
    [string]$CliPath = (Join-Path $PSScriptRoot '../src/ExcelMcp.CLI/bin/Release/net10.0/excelcli')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$sourceRoot = Join-Path $PSScriptRoot '../helper/mac'
$version = (Get-Content -LiteralPath (Join-Path $sourceRoot 'VERSION') -Raw).Trim()
if ($version -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
    throw 'helper/mac/VERSION must contain a stable semantic version.'
}
$output = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($output)
$template = Get-Content -LiteralPath (Join-Path $sourceRoot 'ExcelMcpHelper.bas.template') -Raw
if (-not $template.Contains('@@HELPER_VERSION@@')) { throw 'The helper template has no version placeholder.' }
if ($template.ToCharArray() | Where-Object { [int]$_ -gt 127 }) { throw 'The helper module template must contain ASCII text.' }
$module = Join-Path $output 'ExcelMcpHelper.bas'
[IO.File]::WriteAllText($module, $template.Replace('@@HELPER_VERSION@@', $version), [Text.ASCIIEncoding]::new())
if ($PrepareOnly) {
    Write-Host "Prepared helper $version module: $module"
    return
}
if (-not $IsMacOS) { throw 'Building the helper requires macOS desktop Excel.' }
if (-not $BootstrapPath) { throw 'BootstrapPath is required. Import the prepared module into an Excel-authored .xlsm first.' }
$bootstrap = [IO.Path]::GetFullPath($BootstrapPath)
if (-not [IO.File]::Exists($bootstrap)) { throw 'The Excel-authored helper bootstrap workbook does not exist.' }
if (-not [IO.File]::Exists($CliPath)) { throw 'Build the Release CLI before building the helper.' }
$artifact = Join-Path $output 'ExcelMcpHelper.xlam'
if ([IO.File]::Exists($artifact)) { throw 'The helper artifact already exists; it will not be overwritten.' }
$command = @{ command = 'service.helper-build'; args = @{ workbookPath = $bootstrap; outputPath = $artifact; helperVersion = $version } } |
    ConvertTo-Json -Depth 5 -Compress
$executable = [IO.Path]::GetFullPath($CliPath)
$managedEntry = [IO.Path]::ChangeExtension($executable, '.dll')
$start = if ([IO.File]::Exists($managedEntry)) {
    $managedStart = [Diagnostics.ProcessStartInfo]::new('dotnet')
    $managedStart.ArgumentList.Add($managedEntry)
    $managedStart
} else {
    [Diagnostics.ProcessStartInfo]::new($executable)
}
$start.UseShellExecute = $false
$start.RedirectStandardInput = $true
$start.RedirectStandardOutput = $true
$start.RedirectStandardError = $true
$start.ArgumentList.Add('-q')
$start.ArgumentList.Add('batch')
$process = [Diagnostics.Process]::new()
$process.StartInfo = $start
try {
    if (-not $process.Start()) { throw 'Could not start the helper build client.' }
    $stdout = $process.StandardOutput.ReadToEndAsync()
    $stderr = $process.StandardError.ReadToEndAsync()
    $process.StandardInput.WriteLine($command)
    $process.StandardInput.Close()
    if (-not $process.WaitForExit(150000)) {
        $process.Kill($true)
        $process.WaitForExit()
        throw 'Helper build client timed out. Its Excel outcome may be uncertain; reconcile the artifact and bootstrap before retrying.'
    }
    $response = $stdout.GetAwaiter().GetResult() | ConvertFrom-Json
    if ($process.ExitCode -ne 0 -or -not $response.success) {
        throw "Excel helper build failed: $($response.error) $($stderr.GetAwaiter().GetResult())"
    }
    if ($response.result.helper.version -ne $version) {
        throw "Excel built helper $($response.result.helper.version), but the reviewed source version is $version. Do not publish the artifact."
    }
    if (-not [IO.File]::Exists($artifact)) { throw 'Excel did not create the helper artifact.' }
    $hash = (Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText((Join-Path $output 'SHA256SUMS'), "$hash  ExcelMcpHelper.xlam`n", [Text.ASCIIEncoding]::new())
    Write-Host "Built independently versioned helper ${version}: $artifact"
} finally {
    $process.Dispose()
}

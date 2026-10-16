#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassthroughArgs
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ($null -eq $PassthroughArgs) {
    $PassthroughArgs = @()
}

$npxCommand = Get-Command "npx.cmd" -CommandType Application -ErrorAction SilentlyContinue |
    Select-Object -First 1
$nodeCommand = Get-Command "node.exe" -CommandType Application -ErrorAction SilentlyContinue |
    Select-Object -First 1
$npxCliPath = if ($null -ne $npxCommand) {
    Join-Path (Split-Path -Parent $npxCommand.Source) "node_modules\npm\bin\npx-cli.js"
}

if ($null -ne $nodeCommand -and -not [string]::IsNullOrWhiteSpace($npxCliPath) -and
    (Test-Path -LiteralPath $npxCliPath -PathType Leaf)) {
    $binaryPath = $nodeCommand.Source
    $nativeArguments = @($npxCliPath, "-y", "@sbroenne/excelcli@latest") + @($PassthroughArgs)
} else {
    throw "excel-cli requires Node.js 18 or later with npm/npx available on PATH."
}

$startInfo = [Diagnostics.ProcessStartInfo]::new()
$startInfo.FileName = $binaryPath
foreach ($argument in $nativeArguments) { $startInfo.ArgumentList.Add($argument) }
$startInfo.UseShellExecute = $false
$startInfo.CreateNoWindow = $true
$startInfo.RedirectStandardOutput = $true
$startInfo.RedirectStandardError = $true

$process = [System.Diagnostics.Process]::Start($startInfo)
try {
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    if (-not [string]::IsNullOrEmpty($stderr)) { [Console]::Error.Write($stderr) }
    if (-not [string]::IsNullOrEmpty($stdout)) { Write-Output -NoEnumerate $stdout }
    $exitCode = $process.ExitCode
}
finally { $process.Dispose() }
exit $exitCode

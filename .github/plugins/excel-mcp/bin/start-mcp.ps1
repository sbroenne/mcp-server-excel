[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassthroughArgs
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$downloadScript = Join-Path $PSScriptRoot "download.ps1"
$binaryPath = & $downloadScript -PassThru -Quiet

if ([string]::IsNullOrWhiteSpace($binaryPath) -or -not (Test-Path $binaryPath)) {
    throw "excel-mcp bootstrap did not resolve a usable mcp-excel.exe runtime."
}

& $binaryPath @PassthroughArgs
exit $LASTEXITCODE

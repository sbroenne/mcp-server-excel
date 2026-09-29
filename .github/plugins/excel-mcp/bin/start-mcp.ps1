[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassthroughArgs
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$npx = Get-Command "npx" -CommandType Application, ExternalScript -ErrorAction SilentlyContinue |
    Select-Object -First 1
if ($null -ne $npx) {
    & $npx.Source -y "@sbroenne/mcp-server-excel@latest" @PassthroughArgs
    exit $LASTEXITCODE
}

$downloadScript = Join-Path $PSScriptRoot "download.ps1"
$binaryPath = & $downloadScript -PassThru -Quiet

if ([string]::IsNullOrWhiteSpace($binaryPath) -or -not (Test-Path $binaryPath)) {
    throw "excel-mcp could not run through npx or resolve a fallback mcp-excel.exe runtime."
}

& $binaryPath @PassthroughArgs
exit $LASTEXITCODE

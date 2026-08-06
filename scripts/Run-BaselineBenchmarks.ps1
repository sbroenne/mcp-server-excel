[CmdletBinding()]
param(
    [ValidateSet('quick', 'standard', 'reliable')]
    [string]$Profile = 'standard',

    [string]$Plans = '01,02,03,04,05,06,07,08,09',

    [string]$OutputDirectory,

    [switch]$ShowExcel,

    [int]$MaximumMinutes = 0
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot 'benchmarks\ExcelMcp.Benchmarks\ExcelMcp.Benchmarks.csproj'
$dllPath = Join-Path $repoRoot 'benchmarks\ExcelMcp.Benchmarks\bin\Release\net10.0-windows\Sbroenne.ExcelMcp.Benchmarks.dll'

if ($env:OS -ne 'Windows_NT') {
    throw 'The baseline suite requires Windows and desktop Microsoft Excel.'
}

if (-not [type]::GetTypeFromProgID('Excel.Application')) {
    throw 'Microsoft Excel is not registered on this machine.'
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $OutputDirectory = Join-Path $repoRoot "artifacts\benchmarks\$stamp-$Profile"
}

if ($MaximumMinutes -le 0) {
    $MaximumMinutes = if ($Profile -eq 'reliable') { 180 } else { 90 }
}

$existingExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($existingExcel.Count -gt 0) {
    Write-Warning "Excel is already open ($($existingExcel.Count) process(es)). The harness will not close them, but background activity can add timing noise."
}

$env:ExcelMcpSkipCleanup = 'true'
$env:EXCELMCP_TELEMETRY_OPTOUT = 'true'

dotnet build $projectPath -c Release --nologo /p:ExcelMcpSkipCleanup=true
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$benchmarkArguments = @(
    $dllPath,
    'run',
    '--profile', $Profile,
    '--plans', $Plans,
    '--output', (Resolve-Path -LiteralPath (New-Item -ItemType Directory -Path $OutputDirectory -Force).FullName).Path,
    '--repo', $repoRoot,
    '--maximum-minutes', $MaximumMinutes
)

if ($ShowExcel) {
    $benchmarkArguments += '--show'
}

dotnet @benchmarkArguments
exit $LASTEXITCODE

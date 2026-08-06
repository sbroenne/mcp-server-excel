[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Baseline,

    [Parameter(Mandatory)]
    [string]$Candidate,

    [Parameter(Mandatory)]
    [string]$OutputDirectory,

    [string]$Plans
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$dllPath = Join-Path $repoRoot 'benchmarks\ExcelMcp.Benchmarks\bin\Release\net10.0-windows\Sbroenne.ExcelMcp.Benchmarks.dll'

if (-not (Test-Path -LiteralPath $dllPath)) {
    dotnet build (Join-Path $repoRoot 'benchmarks\ExcelMcp.Benchmarks\ExcelMcp.Benchmarks.csproj') -c Release --nologo /p:ExcelMcpSkipCleanup=true
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

$resolvedOutput = (New-Item -ItemType Directory -Path $OutputDirectory -Force).FullName
$comparisonArguments = @(
    $dllPath,
    'compare',
    '--baseline', (Resolve-Path -LiteralPath $Baseline).Path,
    '--candidate', (Resolve-Path -LiteralPath $Candidate).Path,
    '--output', $resolvedOutput,
    '--repo', $repoRoot
)

if (-not [string]::IsNullOrWhiteSpace($Plans)) {
    $comparisonArguments += @('--plans', $Plans)
}

dotnet @comparisonArguments
exit $LASTEXITCODE

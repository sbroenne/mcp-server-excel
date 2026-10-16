#Requires -Version 7.4
<#
.SYNOPSIS
    Restores the latest validated analytics report from a dedicated data branch.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[^/]+/[^/]+$")]
    [string]$Repository,
    [Parameter(Mandatory = $true)]
    [string]$Branch,
    [Parameter(Mandatory = $true)]
    [string]$ReportPath,
    [Parameter(Mandatory = $true)]
    [string]$RemotePath,
    [scriptblock]$ApiInvoker = {
        param([string[]]$Arguments)
        $output = & gh @Arguments 2>&1
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = @($output) }
    }
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'AnalyticsStorage.ps1')
$resolvedReportPath = [IO.Path]::GetFullPath($ReportPath)
if (-not (Test-Path -LiteralPath $resolvedReportPath -PathType Leaf)) {
    throw "Usage analytics bootstrap file '$resolvedReportPath' does not exist."
}

$stored = Read-AnalyticsFile $Repository $Branch $RemotePath $ApiInvoker `
    'Unable to restore usage analytics' 'GitHub did not return analytics report content.'
if ($stored) {
    $bytes = $stored.Bytes
    $text = [Text.Encoding]::UTF8.GetString($bytes)
    $report = $text | ConvertFrom-Json
    if (@(2, 3) -notcontains $report.schemaVersion -or
        [string]::IsNullOrWhiteSpace([string]$report.interpretation)) {
        throw "Persisted analytics report is incomplete or has an unsupported schema."
    }
    [IO.File]::WriteAllBytes($resolvedReportPath, $bytes)
    $global:LASTEXITCODE = 0
    Write-Host "Restored usage analytics from $Branch."
    return
}
else {
    $global:LASTEXITCODE = 0
    Write-Host "No persisted analytics report exists yet; using the bootstrap report."
    return
}

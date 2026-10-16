#Requires -Version 7.4
<#
.SYNOPSIS
    Persists a validated analytics report through the GitHub Contents API.
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
    [Parameter(Mandatory = $true)]
    [string]$CommitSha,
    [Parameter(Mandatory = $true)]
    [string]$TempPath,
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
    throw "Validated analytics report '$resolvedReportPath' does not exist."
}
$reportText = [IO.File]::ReadAllText($resolvedReportPath)
$report = $reportText | ConvertFrom-Json
if (@(2, 3) -notcontains $report.schemaVersion -or
    [string]::IsNullOrWhiteSpace([string]$report.interpretation)) {
    throw "Analytics report is incomplete or has an unsupported schema."
}

Initialize-AnalyticsBranch $Repository $Branch $CommitSha $ApiInvoker
$existing = Read-AnalyticsFile $Repository $Branch $RemotePath $ApiInvoker `
    'Unable to read persisted analytics report' 'GitHub did not return analytics report content.'
if ($existing) {
    $existingText = [Text.Encoding]::UTF8.GetString($existing.Bytes)
    if ($existingText -eq $reportText) {
        Write-Host "Usage analytics report is already current."
        return
    }
}
Write-AnalyticsFile -Repository $Repository -Branch $Branch -RemotePath $RemotePath -ApiInvoker $ApiInvoker `
    -Text $reportText -Sha $existing.Sha -Message "chore: snapshot usage analytics $([DateTime]::UtcNow.ToString('yyyy-MM-dd'))" `
    -TempPath $TempPath -Prefix 'usage-analytics' -WriteFailure 'Unable to persist usage analytics'

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
$resolvedReportPath = [IO.Path]::GetFullPath($ReportPath)
if (-not (Test-Path -LiteralPath $resolvedReportPath -PathType Leaf)) {
    throw "Usage analytics bootstrap file '$resolvedReportPath' does not exist."
}

$result = & $ApiInvoker -Arguments @(
    "api", ("repos/$Repository/contents/$RemotePath" + "?ref=$Branch")
)
if ($null -eq $result -or
    $null -eq $result.PSObject.Properties["ExitCode"] -or
    $null -eq $result.PSObject.Properties["Output"]) {
    throw "The GitHub API invoker returned an invalid result."
}
$responseText = @($result.Output) -join "`n"
if ($result.ExitCode -eq 0) {
    $response = $responseText | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace([string]$response.content)) {
        throw "GitHub did not return analytics report content."
    }
    $bytes = [Convert]::FromBase64String(($response.content -replace "\s", ""))
    $text = [Text.Encoding]::UTF8.GetString($bytes)
    $report = $text | ConvertFrom-Json
    if ($report.schemaVersion -ne 1 -or
        [string]::IsNullOrWhiteSpace([string]$report.interpretation)) {
        throw "Persisted analytics report is incomplete or has an unsupported schema."
    }
    [IO.File]::WriteAllBytes($resolvedReportPath, $bytes)
    $global:LASTEXITCODE = 0
    Write-Host "Restored usage analytics from $Branch."
    return
}
if ($responseText -match "HTTP 404") {
    $global:LASTEXITCODE = 0
    Write-Host "No persisted analytics report exists yet; using the bootstrap report."
    return
}
throw "Unable to restore usage analytics: $responseText"

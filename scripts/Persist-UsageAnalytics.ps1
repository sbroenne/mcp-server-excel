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
function Invoke-GitHubApi {
    param([string[]]$Arguments)
    $result = & $ApiInvoker -Arguments $Arguments
    if ($null -eq $result -or
        $null -eq $result.PSObject.Properties["ExitCode"] -or
        $null -eq $result.PSObject.Properties["Output"]) {
        throw "The GitHub API invoker returned an invalid result."
    }
    return $result
}

$resolvedReportPath = [IO.Path]::GetFullPath($ReportPath)
if (-not (Test-Path -LiteralPath $resolvedReportPath -PathType Leaf)) {
    throw "Validated analytics report '$resolvedReportPath' does not exist."
}
$reportText = [IO.File]::ReadAllText($resolvedReportPath)
$report = $reportText | ConvertFrom-Json
if ($report.schemaVersion -ne 2 -or
    [string]::IsNullOrWhiteSpace([string]$report.interpretation)) {
    throw "Analytics report is incomplete or has an unsupported schema."
}

$refResult = Invoke-GitHubApi @("api", "repos/$Repository/git/ref/heads/$Branch")
$refText = @($refResult.Output) -join "`n"
if ($refResult.ExitCode -ne 0) {
    if ($refText -notmatch "HTTP 404") {
        throw "Unable to inspect $Branch`: $refText"
    }
    $createResult = Invoke-GitHubApi @(
        "api", "--silent", "--method", "POST",
        "repos/$Repository/git/refs",
        "-f", "ref=refs/heads/$Branch",
        "-f", "sha=$CommitSha"
    )
    if ($createResult.ExitCode -ne 0) {
        throw "Unable to create $Branch`: $(@($createResult.Output) -join "`n")"
    }
}

$contentResult = Invoke-GitHubApi @(
    "api", ("repos/$Repository/contents/$RemotePath" + "?ref=$Branch")
)
$contentText = @($contentResult.Output) -join "`n"
$existingSha = $null
if ($contentResult.ExitCode -eq 0) {
    $existing = $contentText | ConvertFrom-Json
    $existingText = [Text.Encoding]::UTF8.GetString(
        [Convert]::FromBase64String(($existing.content -replace "\s", "")))
    $existingSha = [string]$existing.sha
    if ($existingText -eq $reportText) {
        Write-Host "Usage analytics report is already current."
        return
    }
}
elseif ($contentText -notmatch "HTTP 404") {
    throw "Unable to read persisted analytics report: $contentText"
}

$payload = [ordered]@{
    message = "chore: snapshot usage analytics $([DateTime]::UtcNow.ToString('yyyy-MM-dd'))"
    content = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($reportText))
    branch = $Branch
}
if (-not [string]::IsNullOrWhiteSpace($existingSha)) {
    $payload.sha = $existingSha
}

$payloadPath = Join-Path $TempPath "usage-analytics-$([Guid]::NewGuid().ToString('N')).json"
try {
    [IO.File]::WriteAllText($payloadPath, ($payload | ConvertTo-Json))
    $persistResult = Invoke-GitHubApi @(
        "api", "--silent", "--method", "PUT",
        "repos/$Repository/contents/$RemotePath",
        "--input", $payloadPath
    )
    if ($persistResult.ExitCode -ne 0) {
        throw "Unable to persist usage analytics: $(@($persistResult.Output) -join "`n")"
    }
}
finally {
    Remove-Item -LiteralPath $payloadPath -Force -ErrorAction SilentlyContinue
}

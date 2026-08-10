<#
.SYNOPSIS
    Creates or updates the durable aggregate star-history file through the GitHub Contents API.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[^/]+/[^/]+$")]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [string]$Branch,

    [Parameter(Mandatory = $true)]
    [string]$HistoryPath,

    [Parameter(Mandatory = $true)]
    [string]$RemotePath,

    [Parameter(Mandatory = $true)]
    [string]$CommitSha,

    [Parameter(Mandatory = $true)]
    [string]$TempPath,

    [scriptblock]$ApiInvoker = {
        param([string[]]$Arguments)

        $output = & gh @Arguments 2>&1
        return [pscustomobject]@{
            ExitCode = $LASTEXITCODE
            Output = @($output)
        }
    }
)

$ErrorActionPreference = "Stop"

function Invoke-GitHubApi {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $result = & $ApiInvoker -Arguments $Arguments
    if ($null -eq $result -or
        $null -eq $result.PSObject.Properties["ExitCode"] -or
        $null -eq $result.PSObject.Properties["Output"]) {
        throw "The GitHub API invoker returned an invalid result."
    }

    return $result
}

$resolvedHistoryPath = [IO.Path]::GetFullPath($HistoryPath)
if (-not (Test-Path -LiteralPath $resolvedHistoryPath -PathType Leaf)) {
    throw "Validated star-history aggregate file '$resolvedHistoryPath' does not exist."
}

$updatedText = [IO.File]::ReadAllText($resolvedHistoryPath)
$refApi = "repos/$Repository/git/ref/heads/$Branch"
$refResult = Invoke-GitHubApi -Arguments @("api", $refApi)
$refResponseText = @($refResult.Output) -join "`n"

if ($refResult.ExitCode -ne 0) {
    if ($refResponseText -notmatch "HTTP 404") {
        throw "Unable to inspect $Branch`: $refResponseText"
    }

    $createRefResult = Invoke-GitHubApi -Arguments @(
        "api",
        "--silent",
        "--method", "POST",
        "repos/$Repository/git/refs",
        "-f", "ref=refs/heads/$Branch",
        "-f", "sha=$CommitSha"
    )
    if ($createRefResult.ExitCode -ne 0) {
        throw "Unable to create $Branch`: $(@($createRefResult.Output) -join "`n")"
    }
}

$contentApi = "repos/$Repository/contents/$RemotePath" + "?ref=$Branch"
$contentResult = Invoke-GitHubApi -Arguments @("api", $contentApi)
$contentResponseText = @($contentResult.Output) -join "`n"
$existingSha = $null

if ($contentResult.ExitCode -eq 0) {
    $existing = $contentResponseText | ConvertFrom-Json
    $existingText = [Text.Encoding]::UTF8.GetString(
        [Convert]::FromBase64String(($existing.content -replace "\s", "")))
    $existingSha = [string]$existing.sha

    if ($existingText -eq $updatedText) {
        Write-Host "Today's aggregate snapshot is already current."
        return
    }
}
elseif ($contentResponseText -match "HTTP 404") {
    Write-Host "The aggregate file is absent on $Branch; creating it from the validated snapshot."
}
else {
    throw "Unable to read the persisted aggregate file: $contentResponseText"
}

$payload = [ordered]@{
    message = "chore: snapshot star history $([DateTime]::UtcNow.ToString('yyyy-MM-dd'))"
    content = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($updatedText))
    branch = $Branch
}
if (-not [string]::IsNullOrWhiteSpace($existingSha)) {
    $payload.sha = $existingSha
}

$payloadPath = Join-Path $TempPath "star-history-$([Guid]::NewGuid().ToString('N')).json"
try {
    [IO.File]::WriteAllText($payloadPath, ($payload | ConvertTo-Json))
    $persistResult = Invoke-GitHubApi -Arguments @(
        "api",
        "--silent",
        "--method", "PUT",
        "repos/$Repository/contents/$RemotePath",
        "--input", $payloadPath
    )
    if ($persistResult.ExitCode -ne 0) {
        throw "Unable to persist the daily aggregate snapshot: $(@($persistResult.Output) -join "`n")"
    }
}
finally {
    Remove-Item -LiteralPath $payloadPath -Force -ErrorAction SilentlyContinue
}

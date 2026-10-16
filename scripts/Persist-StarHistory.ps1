#Requires -Version 7.4
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

. (Join-Path $PSScriptRoot 'AnalyticsStorage.ps1')

$resolvedHistoryPath = [IO.Path]::GetFullPath($HistoryPath)
if (-not (Test-Path -LiteralPath $resolvedHistoryPath -PathType Leaf)) {
    throw "Validated star-history aggregate file '$resolvedHistoryPath' does not exist."
}

$updatedText = [IO.File]::ReadAllText($resolvedHistoryPath)
Initialize-AnalyticsBranch $Repository $Branch $CommitSha $ApiInvoker
$existing = Read-AnalyticsFile $Repository $Branch $RemotePath $ApiInvoker `
    'Unable to read the persisted aggregate file' 'GitHub did not return a file content payload for persisted star history.'
if ($existing) {
    $existingText = [Text.Encoding]::UTF8.GetString($existing.Bytes)
    if ($existingText -eq $updatedText) {
        Write-Host "Today's aggregate snapshot is already current."
        return
    }
}
else {
    Write-Host "The aggregate file is absent on $Branch; creating it from the validated snapshot."
}
Write-AnalyticsFile -Repository $Repository -Branch $Branch -RemotePath $RemotePath -ApiInvoker $ApiInvoker `
    -Text $updatedText -Sha $existing.Sha -Message "chore: snapshot star history $([DateTime]::UtcNow.ToString('yyyy-MM-dd'))" `
    -TempPath $TempPath -Prefix 'star-history' -WriteFailure 'Unable to persist the daily aggregate snapshot'

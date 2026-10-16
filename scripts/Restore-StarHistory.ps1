#Requires -Version 7.4
<#
.SYNOPSIS
    Restores durable aggregate star history from the GitHub Contents API.
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
    throw "Star-history bootstrap file '$resolvedHistoryPath' does not exist."
}

$stored = Read-AnalyticsFile $Repository $Branch $RemotePath $ApiInvoker `
    'Unable to restore persisted star history' 'GitHub did not return a file content payload for persisted star history.'
if ($stored) {
    [IO.File]::WriteAllBytes($resolvedHistoryPath, $stored.Bytes)
    $global:LASTEXITCODE = 0
    Write-Host "Restored aggregate snapshots from $Branch."
    return
}

else {
    # GitHub Actions propagates the last native exit code even when this expected 404 is handled.
    $global:LASTEXITCODE = 0
    Write-Host "No persisted branch or aggregate file exists yet; using the exact aggregate bootstrap."
    return
}

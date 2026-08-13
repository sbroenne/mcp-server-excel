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
    throw "Star-history bootstrap file '$resolvedHistoryPath' does not exist."
}

$contentApi = "repos/$Repository/contents/$RemotePath" + "?ref=$Branch"
$contentResult = Invoke-GitHubApi -Arguments @("api", $contentApi)
$responseText = @($contentResult.Output) -join "`n"

if ($contentResult.ExitCode -eq 0) {
    $state = $responseText | ConvertFrom-Json
    $content = [Convert]::FromBase64String(($state.content -replace "\s", ""))
    [IO.File]::WriteAllBytes($resolvedHistoryPath, $content)
    $global:LASTEXITCODE = 0
    Write-Host "Restored aggregate snapshots from $Branch."
    return
}

if ($responseText -match "HTTP 404") {
    # GitHub Actions propagates the last native exit code even when this expected 404 is handled.
    $global:LASTEXITCODE = 0
    Write-Host "No persisted branch or aggregate file exists yet; using the exact aggregate bootstrap."
    return
}

throw "Unable to restore persisted star history: $responseText"

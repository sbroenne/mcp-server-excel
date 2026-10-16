#Requires -Version 7.4

function Invoke-AnalyticsApi {
    param([scriptblock]$ApiInvoker, [string[]]$Arguments)
    $result = & $ApiInvoker -Arguments $Arguments
    if ($null -eq $result -or $null -eq $result.PSObject.Properties['ExitCode'] -or $null -eq $result.PSObject.Properties['Output']) {
        throw 'The GitHub API invoker returned an invalid result.'
    }
    return $result
}

function Initialize-AnalyticsBranch {
    param([string]$Repository, [string]$Branch, [string]$CommitSha, [scriptblock]$ApiInvoker)
    $result = Invoke-AnalyticsApi $ApiInvoker @('api', "repos/$Repository/git/ref/heads/$Branch")
    if ($result.ExitCode -eq 0) { return }
    $text = @($result.Output) -join "`n"
    if ($text -notmatch 'HTTP 404') { throw "Unable to inspect ${Branch}: $text" }
    $created = Invoke-AnalyticsApi $ApiInvoker @('api', '--silent', '--method', 'POST',
        "repos/$Repository/git/refs", '-f', "ref=refs/heads/$Branch", '-f', "sha=$CommitSha")
    if ($created.ExitCode -ne 0) { throw "Unable to create ${Branch}: $(@($created.Output) -join "`n")" }
}

function Read-AnalyticsFile {
    param([string]$Repository, [string]$Branch, [string]$RemotePath, [scriptblock]$ApiInvoker,
        [string]$ReadFailure, [string]$ContentFailure)
    $result = Invoke-AnalyticsApi $ApiInvoker @('api', ("repos/$Repository/contents/$RemotePath" + "?ref=$Branch"))
    $text = @($result.Output) -join "`n"
    if ($result.ExitCode -ne 0) {
        if ($text -match 'HTTP 404') { return $null }
        throw "${ReadFailure}: $text"
    }
    $content = $text | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace([string]$content.content)) { throw $ContentFailure }
    [pscustomobject]@{
        Bytes = [Convert]::FromBase64String(($content.content -replace '\s', ''))
        Sha = [string]$content.sha
    }
}

function Write-AnalyticsFile {
    param([string]$Repository, [string]$Branch, [string]$RemotePath, [scriptblock]$ApiInvoker,
        [string]$Text, [string]$Sha, [string]$Message, [string]$TempPath, [string]$Prefix, [string]$WriteFailure)
    $payload = [ordered]@{ message = $Message; content = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)); branch = $Branch }
    if (-not [string]::IsNullOrWhiteSpace($Sha)) { $payload.sha = $Sha }
    $path = Join-Path $TempPath "$Prefix-$([Guid]::NewGuid().ToString('N')).json"
    $primary = $null
    try {
        [IO.File]::WriteAllText($path, ($payload | ConvertTo-Json))
        $result = Invoke-AnalyticsApi $ApiInvoker @('api', '--silent', '--method', 'PUT', "repos/$Repository/contents/$RemotePath", '--input', $path)
        if ($result.ExitCode -ne 0) { throw "${WriteFailure}: $(@($result.Output) -join "`n")" }
    } catch { $primary = $_; throw }
    finally {
        try { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
        catch {
            if ($primary) { throw [AggregateException]::new('Analytics persistence and temporary-file cleanup failed.', [Exception[]]@($primary.Exception, $_.Exception)) }
            throw
        }
    }
}

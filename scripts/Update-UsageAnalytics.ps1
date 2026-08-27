<#
.SYNOPSIS
    Queries privacy-safe aggregate telemetry for the public usage analytics report.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceId,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [ValidateRange(10, 1000)]
    [int]$MinimumUsers = 10,

    [string]$FixturePath
)

$ErrorActionPreference = "Stop"
$queries = [ordered]@{
    overview = @'
let sessions = AppEvents
| where TimeGenerated > ago(90d)
| summarize HasStart=countif(Name == 'SessionStart') > 0,
            Invocations=countif(Name != 'SessionStart') by SessionId;
let engagement = AppEvents
| where TimeGenerated > ago(90d)
| summarize ActiveDays=dcount(startofday(TimeGenerated)),
            Sessions=dcount(SessionId) by UserId
| summarize RepeatUserRate=round(100.0 * countif(ActiveDays >= 2) / count(), 2),
            MultiSessionRate=round(100.0 * countif(Sessions >= 2) / count(), 2);
sessions
| summarize Sessions=count(),
            ActivatedSessions=countif(Invocations > 0),
            ToolInvocations=sum(Invocations)
| extend ActivationRate=round(100.0 * ActivatedSessions / Sessions, 2)
| extend Users=toscalar(AppEvents | where TimeGenerated > ago(90d) | summarize dcount(UserId))
| extend RepeatUserRate=toscalar(engagement | project RepeatUserRate),
         MultiSessionRate=toscalar(engagement | project MultiSessionRate)
'@
    trend = @'
let current = AppEvents
| where TimeGenerated between (ago(14d) .. now())
| summarize Users=dcount(UserId), Sessions=dcount(SessionId),
            Invocations=countif(Name != 'SessionStart');
let previous = AppEvents
| where TimeGenerated between (ago(28d) .. ago(14d))
| summarize Users=dcount(UserId), Sessions=dcount(SessionId),
            Invocations=countif(Name != 'SessionStart');
current
| extend PreviousUsers=toscalar(previous | project Users),
         PreviousSessions=toscalar(previous | project Sessions),
         PreviousInvocations=toscalar(previous | project Invocations)
| extend UserChangePct=round(100.0 * (Users-PreviousUsers) / PreviousUsers, 2),
         SessionChangePct=round(100.0 * (Sessions-PreviousSessions) / PreviousSessions, 2),
         InvocationChangePct=round(100.0 * (Invocations-PreviousInvocations) / PreviousInvocations, 2)
'@
    operations = @'
AppRequests
| where TimeGenerated > ago(90d)
| summarize Invocations=count(), Users=dcount(UserId),
            SuccessRate=round(100.0 * countif(Success == true) / count(), 2),
            P50Ms=round(percentile(DurationMs, 50), 1),
            P95Ms=round(percentile(DurationMs, 95), 1),
            P99Ms=round(percentile(DurationMs, 99), 1) by Name
| order by Invocations desc
| take 25
'@
    families = @'
let total=toscalar(AppRequests | where TimeGenerated > ago(90d) | count);
AppRequests
| where TimeGenerated > ago(90d)
| extend ToolFamily=tostring(split(Name, '/')[0])
| summarize Invocations=count(), Users=dcount(UserId),
            SuccessRate=round(100.0 * countif(Success == true) / count(), 2),
            P95Ms=round(percentile(DurationMs, 95), 1) by ToolFamily
| extend SharePct=round(100.0 * Invocations / total, 2)
| order by Invocations desc
| take 20
'@
    versions = @'
AppEvents
| where TimeGenerated > ago(14d)
| extend Version=tostring(split(AppVersion, '+')[0])
| summarize Invocations=countif(Name != 'SessionStart'), UserSketch=hll(UserId)
    by Version, SessionId
| summarize Invocations=sum(Invocations), Sessions=count(),
            ActivatedSessions=countif(Invocations > 0),
            UserSketch=hll_merge(UserSketch) by Version
| extend ActivationRate=round(100.0 * ActivatedSessions / Sessions, 2),
         Users=dcount_hll(UserSketch)
| project-away UserSketch
| order by Invocations desc
| take 25
'@
    exceptions = @'
AppExceptions
| where TimeGenerated > ago(90d)
| where tostring(Properties['Sanitized']) == 'true'
    or isnotempty(tostring(Properties['Source']))
| summarize Exceptions=count(), Users=dcount(UserId), Sessions=dcount(SessionId)
| extend Category='background-task-problem'
'@
}

function Invoke-LogAnalyticsQuery {
    param([Parameter(Mandatory = $true)][string]$Query)

    $token = az account get-access-token `
        --resource https://api.loganalytics.io `
        --query accessToken `
        --output tsv `
        --only-show-errors
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($token)) {
        throw "Unable to acquire a Log Analytics access token."
    }

    $body = @{ query = $Query; timespan = "P90D" } | ConvertTo-Json
    $response = Invoke-RestMethod `
        -Method Post `
        -Uri "https://api.loganalytics.azure.com/v1/workspaces/$WorkspaceId/query" `
        -Headers @{ Authorization = "Bearer $token" } `
        -ContentType "application/json" `
        -Body $body
    $table = @($response.tables)[0]
    if ($null -eq $table) {
        throw "Log Analytics returned no result table."
    }

    $columns = @($table.columns.name)
    return @(
        foreach ($row in $table.rows) {
            $item = [ordered]@{}
            for ($index = 0; $index -lt $columns.Count; $index++) {
                $item[$columns[$index]] = $row[$index]
            }
            [pscustomobject]$item
        }
    )
}

function Convert-ToNumber {
    param([object]$Value)
    if ($null -eq $Value -or $Value -isnot [ValueType]) {
        throw "Analytics query returned a non-numeric value."
    }
    return $Value
}

$fixtures = $null
if (-not [string]::IsNullOrWhiteSpace($FixturePath)) {
    $resolvedFixturePath = [IO.Path]::GetFullPath($FixturePath)
    if (-not (Test-Path -LiteralPath $resolvedFixturePath -PathType Leaf)) {
        throw "Analytics fixture '$resolvedFixturePath' does not exist."
    }
    $fixtures = Get-Content -LiteralPath $resolvedFixturePath -Raw | ConvertFrom-Json
}

$results = [ordered]@{}
foreach ($entry in $queries.GetEnumerator()) {
    if ($null -ne $fixtures) {
        $property = $fixtures.PSObject.Properties[$entry.Key]
        if ($null -eq $property) {
            throw "Analytics fixture is missing '$($entry.Key)'."
        }
        $results[$entry.Key] = @($property.Value)
    }
    else {
        $results[$entry.Key] = @(Invoke-LogAnalyticsQuery -Query $entry.Value)
    }
}

$overview = @($results.overview)[0]
$trend = @($results.trend)[0]
if ($null -eq $overview -or $null -eq $trend) {
    throw "Analytics overview and trend queries must each return one row."
}

$report = [ordered]@{
    schemaVersion = 1
    generatedAtUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    windows = [ordered]@{
        reportingDays = 90
        comparisonDays = 14
        versionDays = 14
    }
    privacy = [ordered]@{
        minimumUsersPerDimension = $MinimumUsers
        excluded = @(
            "user identifiers",
            "session identifiers",
            "file hashes",
            "geography",
            "messages",
            "stack traces"
        )
    }
    summary = [ordered]@{
        users = Convert-ToNumber $overview.Users
        sessions = Convert-ToNumber $overview.Sessions
        activatedSessions = Convert-ToNumber $overview.ActivatedSessions
        activationRate = Convert-ToNumber $overview.ActivationRate
        toolInvocations = Convert-ToNumber $overview.ToolInvocations
        repeatUserRate = Convert-ToNumber $overview.RepeatUserRate
        multiSessionRate = Convert-ToNumber $overview.MultiSessionRate
    }
    comparison = [ordered]@{
        currentUsers = Convert-ToNumber $trend.Users
        previousUsers = Convert-ToNumber $trend.PreviousUsers
        userChangePct = Convert-ToNumber $trend.UserChangePct
        currentSessions = Convert-ToNumber $trend.Sessions
        previousSessions = Convert-ToNumber $trend.PreviousSessions
        sessionChangePct = Convert-ToNumber $trend.SessionChangePct
        currentInvocations = Convert-ToNumber $trend.Invocations
        previousInvocations = Convert-ToNumber $trend.PreviousInvocations
        invocationChangePct = Convert-ToNumber $trend.InvocationChangePct
    }
    operations = @(
        $results.operations |
            Where-Object { [int]$_.Users -ge $MinimumUsers } |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Name
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                    successRate = Convert-ToNumber $_.SuccessRate
                    p50Ms = Convert-ToNumber $_.P50Ms
                    p95Ms = Convert-ToNumber $_.P95Ms
                    p99Ms = Convert-ToNumber $_.P99Ms
                }
            }
    )
    toolFamilies = @(
        $results.families |
            Where-Object { [int]$_.Users -ge $MinimumUsers } |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.ToolFamily
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                    sharePct = Convert-ToNumber $_.SharePct
                    successRate = Convert-ToNumber $_.SuccessRate
                    p95Ms = Convert-ToNumber $_.P95Ms
                }
            }
    )
    versions = @(
        $results.versions |
            Where-Object { [int]$_.Users -ge $MinimumUsers } |
            ForEach-Object {
                [ordered]@{
                    version = [string]$_.Version
                    invocations = Convert-ToNumber $_.Invocations
                    sessions = Convert-ToNumber $_.Sessions
                    activatedSessions = Convert-ToNumber $_.ActivatedSessions
                    activationRate = Convert-ToNumber $_.ActivationRate
                    users = Convert-ToNumber $_.Users
                }
            }
    )
    exceptions = @(
        $results.exceptions |
            Where-Object { [int]$_.Users -ge $MinimumUsers } |
            ForEach-Object {
                [ordered]@{
                    category = [string]$_.Category
                    exceptions = Convert-ToNumber $_.Exceptions
                    users = Convert-ToNumber $_.Users
                    sessions = Convert-ToNumber $_.Sessions
                }
            }
    )
}

foreach ($operation in $report.operations) {
    if ($operation.name -notmatch '^[a-z0-9_/-]+$') {
        throw "Analytics contains an unsafe operation dimension."
    }
}
foreach ($family in $report.toolFamilies) {
    if ($family.name -notmatch '^[a-z0-9_-]+$') {
        throw "Analytics contains an unsafe tool-family dimension."
    }
}
foreach ($version in $report.versions) {
    if ($version.version -notmatch '^[0-9A-Za-z.+-]+$') {
        throw "Analytics contains an unsafe version dimension."
    }
}
foreach ($exception in $report.exceptions) {
    if ($exception.category -ne "background-task-problem") {
        throw "Analytics contains an unsafe exception category."
    }
}

$json = $report | ConvertTo-Json -Depth 8
$forbidden = @(
    '"UserId"', '"SessionId"', '"FileSessionId"', '"ClientIP"',
    '"ClientCity"', '"ClientCountryOrRegion"', '"Message"', '"StackTrace"',
    '"ExceptionType"', '"InnerExceptionTypes"', '"FailureSite"',
    'TaskScheduler.UnobservedTaskException', 'AggregateException', 'COMException'
)
foreach ($term in $forbidden) {
    if ($json.Contains($term, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Generated analytics contains forbidden field $term."
    }
}
if ($json -match '[A-Za-z]:\\' -or
    $json -match '\\\\[^\\\s]+\\' -or
    $json -match '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}') {
    throw "Generated analytics contains path or email-shaped content."
}

$resolvedOutputPath = [IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}
[IO.File]::WriteAllText(
    $resolvedOutputPath,
    $json + "`n",
    [Text.UTF8Encoding]::new($false))

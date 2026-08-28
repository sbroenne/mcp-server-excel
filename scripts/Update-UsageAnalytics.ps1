<#
.SYNOPSIS
    Queries privacy-safe aggregate telemetry for the public usage analytics report.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceId,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$FixturePath
)

$ErrorActionPreference = "Stop"
$reliabilitySinceUtc = "2026-08-27T17:05:06Z"
$excludedActions = @("file/open", "file/close")
$queries = [ordered]@{
    overview = @'
let engagement = AppEvents
| where TimeGenerated > ago(90d)
| where Name !in ('SessionStart', 'file/open', 'file/close')
| summarize ActiveDays=dcount(startofday(TimeGenerated)) by UserId
| summarize RepeatUserRate=iif(
    count() == 0,
    0.0,
    round(100.0 * countif(ActiveDays >= 2) / count(), 2));
AppEvents
| where TimeGenerated > ago(90d)
| where Name !in ('SessionStart', 'file/open', 'file/close')
| summarize Users=dcount(UserId),
            ToolInvocations=count()
| extend RepeatUserRate=toscalar(engagement | project RepeatUserRate)
'@
    trend = @'
let current = AppEvents
| where TimeGenerated between (ago(14d) .. now())
| where Name !in ('SessionStart', 'file/open', 'file/close')
| summarize Users=dcount(UserId), Invocations=count();
let previous = AppEvents
| where TimeGenerated between (ago(28d) .. ago(14d))
| where Name !in ('SessionStart', 'file/open', 'file/close')
| summarize Users=dcount(UserId), Invocations=count();
current
| extend PreviousUsers=toscalar(previous | project Users),
         PreviousInvocations=toscalar(previous | project Invocations)
| extend UserChangePct=iif(
             PreviousUsers == 0,
             0.0,
             round(100.0 * (Users-PreviousUsers) / PreviousUsers, 2)),
         InvocationChangePct=iif(
             PreviousInvocations == 0,
             0.0,
             round(100.0 * (Invocations-PreviousInvocations) / PreviousInvocations, 2))
'@
    weekly = @'
AppEvents
| where TimeGenerated >= startofweek(ago(84d))
    and TimeGenerated < startofweek(now())
| where Name !in ('SessionStart', 'file/open', 'file/close')
| summarize Users=dcount(UserId),
            Actions=count()
    by Week=startofweek(TimeGenerated)
| order by Week asc
'@
    versionAdoption = @'
let weeklyUsers = AppRequests
| where TimeGenerated >= startofweek(ago(77d))
    and TimeGenerated <= now()
| where Name !in ('file/open', 'file/close')
| extend Week=startofweek(TimeGenerated),
         Version=tostring(split(AppVersion, '+')[0])
| where isnotempty(Version)
| summarize arg_max(TimeGenerated, Version) by Week, UserId;
let counts = weeklyUsers
| summarize Users=count() by Week, Version;
let popularVersions = counts
| summarize Users=sum(Users) by Version
| top 4 by Users
| project Version;
let newestVersions = counts
| summarize by Version
| extend ParsedVersion=parse_version(Version)
| where isnotnull(ParsedVersion)
| top 4 by ParsedVersion
| project Version;
let displayedVersions = union popularVersions, newestVersions
| distinct Version;
let grouped = counts
| extend Version=iff(Version in (displayedVersions), Version, 'Other')
| summarize Users=sum(Users) by Week, Version;
grouped
| join kind=inner (
    grouped
    | summarize TotalUsers=sum(Users) by Week
) on Week
| extend SharePct=round(100.0 * Users / TotalUsers, 2)
| project Week, Version, Users, SharePct
| order by Week asc, Users desc
'@
    operations = @'
AppRequests
| where TimeGenerated > ago(90d)
| where Name !in ('file/open', 'file/close')
| summarize Invocations=count(), Users=dcount(UserId) by Name
| order by Invocations desc
| take 25
'@
    families = @'
let total=toscalar(
    AppRequests
    | where TimeGenerated > ago(90d)
    | where Name !in ('file/open', 'file/close')
    | count);
AppRequests
| where TimeGenerated > ago(90d)
| where Name !in ('file/open', 'file/close')
| extend ToolFamily=tostring(split(Name, '/')[0])
| summarize Invocations=count(), Users=dcount(UserId) by ToolFamily
| extend SharePct=iif(total == 0, 0.0, round(100.0 * Invocations / total, 2))
| order by Invocations desc
| take 20
'@
    heroFeatures = @'
let total=toscalar(
    AppRequests
    | where TimeGenerated > ago(90d)
    | where Name !in ('file/open', 'file/close')
    | count);
AppRequests
| where TimeGenerated > ago(90d)
| where Name !in ('file/open', 'file/close')
| extend ToolFamily=tostring(split(Name, '/')[0])
| extend HeroFeature=case(
    ToolFamily == 'powerquery', 'power-query',
    ToolFamily in ('datamodel', 'datamodelrel'), 'power-pivot-dax',
    ToolFamily in (
        'pivottable', 'pivottable_field', 'chart', 'chart_config',
        'slicer', 'drawing', 'screenshot', 'conditionalformat'
    ), 'pivottables-charts',
    ToolFamily in (
        'range', 'range_format', 'range_edit', 'range_link',
        'table', 'calculation_mode'
    ), 'tables-ranges',
    ToolFamily == 'vba', 'vba',
    ToolFamily in (
        'worksheet', 'worksheet_style', 'workbook', 'file',
        'namedrange', 'connection', 'querytable'
    ), 'worksheets-connections',
    ToolFamily == 'window', 'agent-mode',
    ToolFamily == 'pythoninexcel', 'python-in-excel',
    'other'
)
| summarize Invocations=count(), Users=dcount(UserId) by HeroFeature
| extend SharePct=iif(total == 0, 0.0, round(100.0 * Invocations / total, 2))
| order by Invocations desc
'@
    reliability = @"
AppRequests
| where TimeGenerated >= datetime($reliabilitySinceUtc)
| where Name !in ('file/open', 'file/close')
| extend Version=tostring(split(AppVersion, '+')[0])
| where parse_version(Version) >= parse_version('2.0.3')
| summarize Actions=count(),
            Errors=countif(Success != true),
            Users=dcount(UserId) by Name
| where Errors > 0
| extend ErrorRate=round(100.0 * Errors / Actions, 2)
| order by Errors desc
| take 20
"@
    versionReliability = @"
AppRequests
| where TimeGenerated >= datetime($reliabilitySinceUtc)
| where Name !in ('file/open', 'file/close')
| extend Version=tostring(split(AppVersion, '+')[0])
| where parse_version(Version) >= parse_version('2.0.3')
| summarize Actions=count(),
            Errors=countif(Success != true),
            Users=dcount(UserId) by Version
| extend ErrorRate=round(100.0 * Errors / Actions, 2)
| order by Actions desc
| take 25
"@
    exceptions = @"
AppExceptions
| where TimeGenerated >= datetime($reliabilitySinceUtc)
| where tostring(Properties['Sanitized']) == 'true'
| summarize Exceptions=count(), Users=dcount(UserId), Sessions=dcount(SessionId)
| where Exceptions > 0
| extend Category='background-task-problem'
"@
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
    $numericTypes = @(
        [int], [long], [double], [decimal], [single]
    )
    if ($null -eq $Value -or
        -not ($numericTypes | Where-Object { $_.IsInstanceOfType($Value) })) {
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
        trendWeeks = 12
        reliabilitySinceUtc = $reliabilitySinceUtc
    }
    privacy = [ordered]@{
        excluded = @(
            "user identifiers",
            "session identifiers",
            "names and account details",
            "locations",
            "workbook contents",
            "cell values and formulas",
            "prompts and messages",
            "file names and paths",
            "error messages and stack traces"
        )
    }
    summary = [ordered]@{
        users = Convert-ToNumber $overview.Users
        toolInvocations = Convert-ToNumber $overview.ToolInvocations
        repeatUserRate = Convert-ToNumber $overview.RepeatUserRate
    }
    comparison = [ordered]@{
        currentUsers = Convert-ToNumber $trend.Users
        previousUsers = Convert-ToNumber $trend.PreviousUsers
        userChangePct = Convert-ToNumber $trend.UserChangePct
        currentInvocations = Convert-ToNumber $trend.Invocations
        previousInvocations = Convert-ToNumber $trend.PreviousInvocations
        invocationChangePct = Convert-ToNumber $trend.InvocationChangePct
    }
    weekly = @(
        $results.weekly |
            ForEach-Object {
                [ordered]@{
                    week = ([DateTime]$_.Week).ToString("yyyy-MM-dd")
                    users = Convert-ToNumber $_.Users
                    actions = Convert-ToNumber $_.Actions
                }
            }
    )
    versionAdoption = @(
        $results.versionAdoption |
            ForEach-Object {
                [ordered]@{
                    week = ([DateTime]$_.Week).ToString("yyyy-MM-dd")
                    version = [string]$_.Version
                    users = Convert-ToNumber $_.Users
                    sharePct = Convert-ToNumber $_.SharePct
                }
            }
    )
    operations = @(
        $results.operations |
            Where-Object { $excludedActions -notcontains [string]$_.Name } |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Name
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                }
            }
    )
    toolFamilies = @(
        $results.families |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.ToolFamily
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                    sharePct = Convert-ToNumber $_.SharePct
                }
            }
    )
    heroFeatures = @(
        $results.heroFeatures |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.HeroFeature
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                    sharePct = Convert-ToNumber $_.SharePct
                }
            }
    )
    reliability = @(
        $results.reliability |
            Where-Object { $excludedActions -notcontains [string]$_.Name } |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Name
                    actions = Convert-ToNumber $_.Actions
                    errors = Convert-ToNumber $_.Errors
                    errorRate = Convert-ToNumber $_.ErrorRate
                    users = Convert-ToNumber $_.Users
                }
            }
    )
    versionReliability = @(
        $results.versionReliability |
            ForEach-Object {
                [ordered]@{
                    version = [string]$_.Version
                    actions = Convert-ToNumber $_.Actions
                    errors = Convert-ToNumber $_.Errors
                    errorRate = Convert-ToNumber $_.ErrorRate
                    users = Convert-ToNumber $_.Users
                }
            }
    )
    exceptions = @(
        $results.exceptions |
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
foreach ($reliabilityItem in $report.reliability) {
    if ($reliabilityItem.name -notmatch '^[a-z0-9_/-]+$') {
        throw "Analytics contains an unsafe reliability dimension."
    }
}
foreach ($family in $report.toolFamilies) {
    if ($family.name -notmatch '^[a-z0-9_-]+$') {
        throw "Analytics contains an unsafe tool-family dimension."
    }
}
foreach ($feature in $report.heroFeatures) {
    if ($feature.name -notmatch '^[a-z0-9-]+$') {
        throw "Analytics contains an unsafe homepage-feature dimension."
    }
}
foreach ($version in $report.versionReliability) {
    if ($version.version -notmatch '^[0-9A-Za-z.+-]+$') {
        throw "Analytics contains an unsafe version dimension."
    }
}
foreach ($week in $report.weekly) {
    if ($week.week -notmatch '^\d{4}-\d{2}-\d{2}$') {
        throw "Analytics contains an unsafe weekly date."
    }
}
foreach ($release in $report.versionAdoption) {
    if ($release.week -notmatch '^\d{4}-\d{2}-\d{2}$' -or
        $release.version -notmatch '^(?:[0-9A-Za-z.+-]+|Other)$') {
        throw "Analytics contains an unsafe release-adoption dimension."
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
if ($json -match '"file/(?:open|close)"') {
    throw "Generated analytics contains excluded workbook lifecycle actions."
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

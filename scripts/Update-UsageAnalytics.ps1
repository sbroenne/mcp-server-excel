<#
.SYNOPSIS
    Queries privacy-safe aggregate telemetry for the public usage analytics report.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceId,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$FixturePath,

    [string]$WeightsPath = (Join-Path $PSScriptRoot "../.github/usage-analytics-weights.json"),

    # The last published report. Its download snapshots are carried forward, because
    # GitHub releases and the VS Code Marketplace only publish running totals.
    [string]$PreviousReportPath,

    [string]$Repository = "sbroenne/mcp-server-excel"
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "UsageAnalyticsWeights.ps1")
$weights = Read-UsageAnalyticsWeights -Path $WeightsPath
# excelcli began reporting usage, with the EntryPoint label, in release 2.0.12.
$entryPointSinceUtc = "2026-09-28T00:00:00Z"
$entryPointMinimumUsers = 10
$entryPoints = @("cli", "mcp-server")
$excludedActions = $weights.ExcludedActions
$habitDays = 30
$habitMinimumUsers = 10
$sessionSizes = @("1", "2-10", "11-50", "51-200", "201+")
$advancedFeatures = @("power-query", "power-pivot-dax", "pivottables-charts", "vba")
$advancedFeatureList = ($advancedFeatures | ForEach-Object { "'$_'" }) -join ", "
$weekdayNames = @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
$weekdayWeeks = 8
$queries = [ordered]@{
    overview = @'
let engagement = normalizedEvents
| where TimeGenerated > ago(90d)
| summarize ActiveDays=dcount(startofday(TimeGenerated)) by UserId
| summarize RepeatUserRate=iif(
    count() == 0,
    0.0,
    round(100.0 * countif(ActiveDays >= 2) / count(), 2));
normalizedEvents
| where TimeGenerated > ago(90d)
| summarize Users=dcount(UserId),
            ToolInvocations=count()
| extend RepeatUserRate=toscalar(engagement | project RepeatUserRate)
'@
    trend = @'
let current = normalizedEvents
| where TimeGenerated between (ago(14d) .. now())
| summarize Users=dcount(UserId), Invocations=count();
let previous = normalizedEvents
| where TimeGenerated between (ago(28d) .. ago(14d))
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
normalizedEvents
| where TimeGenerated >= startofweek(ago(84d))
    and TimeGenerated < startofweek(now())
| summarize Users=dcount(UserId),
            Actions=count()
    by Week=startofweek(TimeGenerated)
| order by Week asc
'@
    versionAdoption = @'
let weeklyUsers = normalizedRequests
| where TimeGenerated >= startofweek(ago(77d))
    and TimeGenerated <= now()
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
normalizedRequests
| where TimeGenerated > ago(90d)
| summarize Invocations=count(), Users=dcount(UserId) by Name
| order by Invocations desc
| take 25
'@
    families = @'
let total=toscalar(
    normalizedRequests
    | where TimeGenerated > ago(90d)
    | count);
normalizedRequests
| where TimeGenerated > ago(90d)
| extend ToolFamily=tostring(split(Name, '/')[0])
| summarize Invocations=count(), Users=dcount(UserId) by ToolFamily
| extend SharePct=iif(total == 0, 0.0, round(100.0 * Invocations / total, 2))
| order by Invocations desc
| take 20
'@
    heroFeatures = @'
let total=toscalar(
    normalizedRequests
    | where TimeGenerated > ago(90d)
    | count);
normalizedRequests
| where TimeGenerated > ago(90d)
| summarize Invocations=count(), Users=dcount(UserId) by HeroFeature=Feature
| extend SharePct=iif(total == 0, 0.0, round(100.0 * Invocations / total, 2))
| order by Invocations desc
'@
    actionCounts = @'
normalizedRequests
| where TimeGenerated > ago(90d)
| summarize Actions=count(), Users=dcount(UserId) by Name
'@
    weeklyActions = @'
normalizedRequests
| where TimeGenerated >= startofweek(ago(84d))
    and TimeGenerated < startofweek(now())
| summarize Actions=count() by Week=startofweek(TimeGenerated), Name
'@
    comparisonActions = @'
normalizedRequests
| where TimeGenerated between (ago(28d) .. now())
| summarize CurrentActions=countif(TimeGenerated >= ago(14d)),
            PreviousActions=countif(TimeGenerated < ago(14d))
    by Name
'@
    heavyWork = @'
normalizedRequests
| where TimeGenerated > ago(90d)
| summarize Users=dcount(UserId),
            HeavyUsers=dcountif(UserId, Name in (heavyNames))
'@
    entryPoints = @"
normalizedRequests
| where TimeGenerated >= max_of(datetime($entryPointSinceUtc), ago(90d))
| where EntryPoint in ('cli', 'mcp-server')
| summarize Users=dcount(UserId), Actions=count() by EntryPoint
"@
    entryPointActions = @"
normalizedRequests
| where TimeGenerated >= max_of(datetime($entryPointSinceUtc), ago(90d))
| where EntryPoint in ('cli', 'mcp-server')
| summarize Actions=count() by EntryPoint, Name
"@
    # Habits use AI assistant sessions only: each excelcli command runs as its own process,
    # so a command line "session" is always a single command.
    assistantSessions = @"
normalizedRequests
| where TimeGenerated > ago(${habitDays}d)
| where EntryPoint == 'mcp-server' and isnotempty(SessionId)
| summarize Actions=count(), Features=dcount(Feature) by SessionId
| extend Size=case(
    Actions == 1, '1',
    Actions <= 10, '2-10',
    Actions <= 50, '11-50',
    Actions <= 200, '51-200',
    '201+')
| summarize Sessions=count(),
            Actions=sum(Actions),
            MultiFeatureSessions=countif(Features >= 2)
    by Size
"@
    assistantSessionMedian = @"
normalizedRequests
| where TimeGenerated > ago(${habitDays}d)
| where EntryPoint == 'mcp-server' and isnotempty(SessionId)
| summarize Actions=count() by SessionId
| summarize MedianActions=percentile(Actions, 50)
"@
    featurePairs = @"
normalizedRequests
| where TimeGenerated > ago(${habitDays}d)
| where EntryPoint == 'mcp-server' and isnotempty(SessionId)
| summarize Features=make_set(Feature), UserId=take_any(UserId) by SessionId
| mv-expand First=Features to typeof(string)
| mv-expand Second=Features to typeof(string)
| where strcmp(First, Second) < 0
| summarize Sessions=count(), Users=dcount(UserId) by First, Second
| where Users >= $habitMinimumUsers
| top 8 by Sessions desc
"@
    returningUsers = @'
let newUsers = normalizedRequests
| summarize FirstSeen=min(TimeGenerated) by UserId
| where FirstSeen between (ago(84d) .. ago(28d));
newUsers
| join kind=inner (normalizedRequests | project UserId, TimeGenerated) on UserId
| summarize AfterWeek=countif(TimeGenerated >= FirstSeen + 7d),
            AfterThreeWeeks=countif(TimeGenerated >= FirstSeen + 21d)
    by UserId
| summarize NewUsers=count(),
            ReturnedAfterWeek=countif(AfterWeek > 0),
            ReturnedAfterThreeWeeks=countif(AfterThreeWeeks > 0)
'@
    featureWait = @"
normalizedRequests
| where TimeGenerated > ago(${habitDays}d)
| summarize Actions=count(),
            Users=dcount(UserId),
            TypicalMs=percentile(DurationMs, 50),
            SlowMs=percentile(DurationMs, 90)
    by Feature
| where Users >= $habitMinimumUsers
| order by TypicalMs desc
"@
    weekdays = @"
normalizedRequests
| where TimeGenerated >= startofweek(ago($($weekdayWeeks * 7)d))
    and TimeGenerated < startofweek(now())
| summarize Actions=count(), Users=dcount(UserId)
    by Day=toint(dayofweek(TimeGenerated) / 1d)
| order by Day asc
"@
    firstAdvancedUse = @"
let firstUse = normalizedRequests
| summarize FirstSeen=min(TimeGenerated) by UserId;
normalizedRequests
| where Feature in ($advancedFeatureList)
| summarize FirstFeatureUse=min(TimeGenerated) by UserId, Feature
| join kind=inner firstUse on UserId
| extend Delay=FirstFeatureUse - FirstSeen
| summarize Users=count(),
            FirstDay=countif(Delay < 1d),
            FirstWeek=countif(Delay >= 1d and Delay < 7d),
            Later=countif(Delay >= 7d)
    by Feature
| where Users >= $habitMinimumUsers
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

$npmPackages = [ordered]@{
    "npm-mcp-server" = "@sbroenne/mcp-server-excel"
    "npm-cli" = "@sbroenne/excelcli"
}
# NuGet is left out: its counts are dominated by automated mirrors (roughly the same
# few hundred downloads for every version), and few people install it from there.
$downloadChannelLabels = [ordered]@{
    "npm-mcp-server" = "npm: MCP Server"
    "npm-cli" = "npm: command line"
    "github-releases" = "GitHub release files"
    "vscode" = "VS Code extension installs"
}
# npm publishes daily history. The other sources only publish running totals, so the
# report keeps one dated snapshot per run and calculates the gain between snapshots.
$snapshotChannels = @("github-releases", "vscode")
# Checksums and release metadata are fetched by automation, not people.
$releaseFilePattern = '\.(?:zip|vsix|mcpb)$'
$maxDownloadSnapshots = 104
$downloadReleaseCount = 10
$npmWeekCount = 12

function ConvertTo-UtcDate {
    param([Parameter(Mandatory = $true)][object]$Value)
    if ($Value -is [DateTime]) {
        $date = if ($Value.Kind -eq [DateTimeKind]::Local) { $Value.ToUniversalTime() } else { $Value }
        return $date.Date
    }
    return [DateTimeOffset]::Parse(
        [string]$Value,
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal).UtcDateTime.Date
}

function Get-DownloadSources {
    $sources = [ordered]@{
        npm = [ordered]@{}
        releases = @()
        vscodeInstalls = $null
    }
    foreach ($entry in $npmPackages.GetEnumerator()) {
        $response = Invoke-RestMethod -Uri "https://api.npmjs.org/downloads/range/last-year/$($entry.Value)"
        $sources.npm[$entry.Key] = @($response.downloads)
    }

    $headers = @{ Accept = "application/vnd.github+json"; "X-GitHub-Api-Version" = "2022-11-28" }
    $token = if (-not [string]::IsNullOrWhiteSpace($env:GH_TOKEN)) { $env:GH_TOKEN } else { $env:GITHUB_TOKEN }
    if (-not [string]::IsNullOrWhiteSpace($token)) {
        $headers.Authorization = "Bearer $token"
    }
    $releases = [Collections.Generic.List[object]]::new()
    for ($page = 1; $page -le 20; $page++) {
        $response = Invoke-RestMethod -Headers $headers `
            -Uri "https://api.github.com/repos/$Repository/releases?per_page=100&page=$page"
        $batch = @($response)
        foreach ($release in $batch) {
            $releases.Add([pscustomobject]@{
                tag = $release.tag_name
                publishedAt = $release.published_at
                draft = $release.draft
                assets = @($release.assets | ForEach-Object {
                    [pscustomobject]@{ name = $_.name; downloads = $_.download_count }
                })
            })
        }
        if ($batch.Count -lt 100) {
            break
        }
    }
    if ($releases.Count -eq 0) {
        throw "GitHub returned no releases for '$Repository'."
    }
    $sources.releases = $releases.ToArray()

    $body = @{
        filters = @(@{ criteria = @(@{ filterType = 7; value = "sbroenne.excel-mcp" }) })
        flags = 914
    } | ConvertTo-Json -Depth 6
    $response = Invoke-RestMethod `
        -Method Post `
        -Uri "https://marketplace.visualstudio.com/_apis/public/gallery/extensionquery" `
        -Headers @{ Accept = "application/json;api-version=7.2-preview.1" } `
        -ContentType "application/json" `
        -Body $body
    $extension = @(@($response.results)[0].extensions)[0]
    $installs = @($extension.statistics | Where-Object { $_.statisticName -eq "install" })[0]
    if ($null -eq $installs) {
        throw "The VS Code Marketplace did not return an install count."
    }
    $sources.vscodeInstalls = $installs.value
    # Match the fixture shape, so both paths are read the same way.
    return $sources | ConvertTo-Json -Depth 8 | ConvertFrom-Json
}

function Get-PreviousDownloadSnapshots {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @()
    }
    $resolvedPath = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        throw "Previous analytics report '$resolvedPath' does not exist."
    }
    $previous = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
    $downloadsProperty = $previous.PSObject.Properties["downloads"]
    if ($null -eq $downloadsProperty) {
        return @()
    }
    return @(
        foreach ($snapshot in @($downloadsProperty.Value.snapshots)) {
            $totals = [ordered]@{}
            foreach ($channel in $snapshotChannels) {
                $property = $snapshot.totals.PSObject.Properties[$channel]
                if ($null -eq $property) {
                    throw "Previous download snapshot is missing '$channel'."
                }
                $totals[$channel] = Convert-ToNumber $property.Value
            }
            [ordered]@{
                date = (ConvertTo-UtcDate $snapshot.date).ToString("yyyy-MM-dd")
                totals = $totals
            }
        }
    )
}

function New-DownloadReport {
    param(
        [Parameter(Mandatory = $true)][object]$Sources,
        [object[]]$PreviousSnapshots = @()
    )

    $collected = [DateTime]::UtcNow
    $totals = [ordered]@{}
    $npmDays = @{}
    foreach ($channel in $npmPackages.Keys) {
        $series = $Sources.npm.PSObject.Properties[$channel]
        if ($null -eq $series) {
            throw "npm download history is missing '$channel'."
        }
        $sum = [long]0
        foreach ($day in @($series.Value)) {
            $count = [long](Convert-ToNumber $day.downloads)
            $sum += $count
            $date = ConvertTo-UtcDate $day.day
            if (-not $npmDays.ContainsKey($date)) {
                $npmDays[$date] = [ordered]@{}
            }
            $npmDays[$date][$channel] = $count
        }
        $totals[$channel] = $sum
    }

    $releaseRows = @(
        foreach ($release in @($Sources.releases | Where-Object { -not $_.draft })) {
            $version = ([string]$release.tag) -replace '^v', ''
            if ($version -notmatch '^[0-9A-Za-z.+-]+$') {
                throw "Download data contains an unsafe release version."
            }
            $downloads = [long]0
            foreach ($asset in @($release.assets)) {
                if ([string]$asset.name -match $releaseFilePattern) {
                    $downloads += [long](Convert-ToNumber $asset.downloads)
                }
            }
            [pscustomobject]@{
                version = $version
                published = ConvertTo-UtcDate $release.publishedAt
                downloads = $downloads
            }
        }
    )
    $totals["github-releases"] = [long](($releaseRows | Measure-Object downloads -Sum).Sum)
    $totals["vscode"] = [long](Convert-ToNumber $Sources.vscodeInstalls)

    # Weeks start on Sunday to match the usage charts; only full weeks are shown.
    $npmWeekly = @()
    if ($npmDays.Count -gt 0) {
        $lastDay = $npmDays.Keys | Sort-Object | Select-Object -Last 1
        $firstActiveDay = $npmDays.Keys |
            Where-Object { ($npmDays[$_].Values | Measure-Object -Sum).Sum -gt 0 } |
            Sort-Object |
            Select-Object -First 1
        if ($null -ne $firstActiveDay) {
            $weeks = [ordered]@{}
            foreach ($date in ($npmDays.Keys | Sort-Object)) {
                $weekStart = $date.AddDays(-[int]$date.DayOfWeek)
                if ($weekStart.AddDays(6) -gt $lastDay -or $weekStart.AddDays(6) -lt $firstActiveDay) {
                    continue
                }
                $key = $weekStart.ToString("yyyy-MM-dd")
                if (-not $weeks.Contains($key)) {
                    $weeks[$key] = [ordered]@{ week = $key; mcpServer = [long]0; cli = [long]0; total = [long]0 }
                }
                $mcpServer = [long]$npmDays[$date]["npm-mcp-server"]
                $cli = [long]$npmDays[$date]["npm-cli"]
                $weeks[$key].mcpServer += $mcpServer
                $weeks[$key].cli += $cli
                $weeks[$key].total += $mcpServer + $cli
            }
            $npmWeekly = @($weeks.Values | Select-Object -Last $npmWeekCount)
        }
    }

    $today = $collected.ToString("yyyy-MM-dd")
    $snapshotTotals = [ordered]@{}
    foreach ($channel in $snapshotChannels) {
        $snapshotTotals[$channel] = $totals[$channel]
    }
    $snapshots = @(
        @($PreviousSnapshots | Where-Object { $_.date -lt $today }) +
            @([ordered]@{ date = $today; totals = $snapshotTotals }) |
            Sort-Object { $_.date } |
            Select-Object -Last $maxDownloadSnapshots
    )
    $gains = @(
        for ($index = 1; $index -lt $snapshots.Count; $index++) {
            $before = $snapshots[$index - 1]
            $after = $snapshots[$index]
            $channelGains = [ordered]@{}
            foreach ($channel in $snapshotChannels) {
                $channelGains[$channel] = [long]$after.totals[$channel] - [long]$before.totals[$channel]
            }
            [ordered]@{
                week = $before.date
                days = ([DateTime]$after.date - [DateTime]$before.date).Days
                total = [long](($channelGains.Values | Measure-Object -Sum).Sum)
                channels = $channelGains
            }
        }
    )

    return [ordered]@{
        collectedUtc = $collected.ToString("yyyy-MM-ddTHH:mm:ssZ")
        channels = @(
            foreach ($entry in $downloadChannelLabels.GetEnumerator()) {
                [ordered]@{ key = $entry.Key; label = $entry.Value; total = $totals[$entry.Key] }
            }
        )
        npmWeekly = $npmWeekly
        releases = @(
            $releaseRows |
                Sort-Object published, version -Descending |
                Select-Object -First $downloadReleaseCount |
                ForEach-Object {
                    [ordered]@{
                        version = $_.version
                        published = $_.published.ToString("yyyy-MM-dd")
                        downloads = $_.downloads
                    }
                }
        )
        snapshotChannels = $snapshotChannels
        snapshots = $snapshots
        weeklyGains = $gains
    }
}
if (-not [string]::IsNullOrWhiteSpace($FixturePath)) {
    $resolvedFixturePath = [IO.Path]::GetFullPath($FixturePath)
    if (-not (Test-Path -LiteralPath $resolvedFixturePath -PathType Leaf)) {
        throw "Analytics fixture '$resolvedFixturePath' does not exist."
    }
    $fixtures = Get-Content -LiteralPath $resolvedFixturePath -Raw | ConvertFrom-Json
}

$queryPrelude = New-UsageAnalyticsQueryPrelude -Weights $weights
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
        $results[$entry.Key] = @(Invoke-LogAnalyticsQuery -Query ($queryPrelude + $entry.Value))
    }
}

$downloadSources = if ($null -ne $fixtures) {
    $property = $fixtures.PSObject.Properties["downloadSources"]
    if ($null -eq $property) {
        throw "Analytics fixture is missing 'downloadSources'."
    }
    $property.Value
}
else {
    Get-DownloadSources
}
$downloads = New-DownloadReport `
    -Sources $downloadSources `
    -PreviousSnapshots @(Get-PreviousDownloadSnapshots -Path $PreviousReportPath)

$overview = @($results.overview)[0]
$trend = @($results.trend)[0]
if ($null -eq $overview -or $null -eq $trend) {
    throw "Analytics overview and trend queries must each return one row."
}

function Get-Percent {
    param([double]$Part, [double]$Total)
    if ($Total -eq 0) {
        return 0.0
    }
    return [Math]::Round(100.0 * $Part / $Total, 2)
}

function Get-WorkUnits {
    param([string]$Name, [object]$Actions)
    if (-not $weights.WeightByName.ContainsKey($Name)) {
        return $null
    }
    return [long]$weights.WeightByName[$Name] * (Convert-ToNumber $Actions)
}

# Work units multiply each action count by its hand-picked level from the weights file.
# Actions without a level (for example, renamed actions from old releases) stay visible as
# unweighted instead of being guessed.
$actionRows = @(
    $results.actionCounts |
        Where-Object { $excludedActions -notcontains [string]$_.Name } |
        ForEach-Object {
            $name = [string]$_.Name
            [pscustomobject]@{
                Name = $name
                Tool = $name.Split('/')[0]
                Feature = Get-UsageAnalyticsFeature -Weights $weights -Name $name
                Actions = Convert-ToNumber $_.Actions
                Users = Convert-ToNumber $_.Users
                WorkUnits = Get-WorkUnits -Name $name -Actions $_.Actions
            }
        }
)
$weightedRows = @($actionRows | Where-Object { $null -ne $_.WorkUnits })
$unweightedRows = @($actionRows | Where-Object { $null -eq $_.WorkUnits })
$totalWorkUnits = [long](($weightedRows | Measure-Object WorkUnits -Sum).Sum)
$workByFeature = @{}
$workByFamily = @{}
foreach ($row in $weightedRows) {
    $workByFeature[$row.Feature] = [long]$workByFeature[$row.Feature] + $row.WorkUnits
    $workByFamily[$row.Tool] = [long]$workByFamily[$row.Tool] + $row.WorkUnits
}

$workByWeek = @{}
foreach ($row in $results.weeklyActions) {
    $work = Get-WorkUnits -Name ([string]$row.Name) -Actions $row.Actions
    if ($null -ne $work -and $excludedActions -notcontains [string]$row.Name) {
        $week = ([DateTime]$row.Week).ToString("yyyy-MM-dd")
        $workByWeek[$week] = [long]$workByWeek[$week] + $work
    }
}

$currentWorkUnits = [long]0
$previousWorkUnits = [long]0
foreach ($row in $results.comparisonActions) {
    if ($excludedActions -contains [string]$row.Name) {
        continue
    }
    $current = Get-WorkUnits -Name ([string]$row.Name) -Actions $row.CurrentActions
    $previous = Get-WorkUnits -Name ([string]$row.Name) -Actions $row.PreviousActions
    if ($null -ne $current) {
        $currentWorkUnits += $current
        $previousWorkUnits += $previous
    }
}

$heavyWork = @($results.heavyWork)[0]
$heavyWorkUsers = if ($null -eq $heavyWork) { 0 } else { Convert-ToNumber $heavyWork.Users }
$heavyWorkHeavyUsers = if ($null -eq $heavyWork) { 0 } else { Convert-ToNumber $heavyWork.HeavyUsers }

$entryPointWindowStart = [DateTime]::Parse(
    $entryPointSinceUtc,
    [Globalization.CultureInfo]::InvariantCulture,
    [Globalization.DateTimeStyles]::AdjustToUniversal)
$reportingWindowStart = [DateTime]::UtcNow.Date.AddDays(-90)
if ($reportingWindowStart -gt $entryPointWindowStart) {
    $entryPointWindowStart = $reportingWindowStart
}
$entryPointSummaries = [Collections.Generic.List[object]]::new()
$entryPointFeatureRows = [Collections.Generic.List[object]]::new()
$entryPointOperationRows = [Collections.Generic.List[object]]::new()
foreach ($row in $results.entryPoints) {
    $name = [string]$row.EntryPoint
    if ($entryPoints -notcontains $name) {
        throw "Analytics contains an unsafe entry point."
    }
    $users = Convert-ToNumber $row.Users
    if ($users -lt $entryPointMinimumUsers) {
        $entryPointSummaries.Add([ordered]@{ name = $name; enoughData = $false })
        continue
    }

    $rows = @(
        $results.entryPointActions |
            Where-Object {
                [string]$_.EntryPoint -eq $name -and
                $excludedActions -notcontains [string]$_.Name
            } |
            ForEach-Object {
                $actionName = [string]$_.Name
                [pscustomobject]@{
                    Name = $actionName
                    Feature = Get-UsageAnalyticsFeature -Weights $weights -Name $actionName
                    Actions = Convert-ToNumber $_.Actions
                    WorkUnits = Get-WorkUnits -Name $actionName -Actions $_.Actions
                }
            }
    )
    $actions = [long](($rows | Measure-Object Actions -Sum).Sum)
    $work = [long](($rows | Where-Object { $null -ne $_.WorkUnits } | Measure-Object WorkUnits -Sum).Sum)
    $entryPointSummaries.Add([ordered]@{
        name = $name
        enoughData = $true
        users = $users
        actions = $actions
        workUnits = $work
        actionsPerUser = [Math]::Round($actions / $users, 1)
        workUnitsPerUser = [Math]::Round($work / $users, 1)
    })
    foreach ($group in ($rows | Group-Object Feature)) {
        $featureActions = [long](($group.Group | Measure-Object Actions -Sum).Sum)
        $featureWork = [long](($group.Group | Where-Object { $null -ne $_.WorkUnits } |
            Measure-Object WorkUnits -Sum).Sum)
        $entryPointFeatureRows.Add([ordered]@{
            entryPoint = $name
            name = $group.Name
            actions = $featureActions
            actionSharePct = Get-Percent $featureActions $actions
            workUnits = $featureWork
            workSharePct = Get-Percent $featureWork $work
        })
    }
    foreach ($operation in ($rows | Sort-Object -Property @{ Expression = "Actions"; Descending = $true }, Name |
        Select-Object -First 10)) {
        $entryPointOperationRows.Add([ordered]@{
            entryPoint = $name
            name = $operation.Name
            actions = $operation.Actions
            workUnits = if ($null -eq $operation.WorkUnits) { 0 } else { $operation.WorkUnits }
        })
    }
}

$sessionRows = @(
    $results.assistantSessions |
        ForEach-Object {
            [pscustomobject]@{
                Size = [string]$_.Size
                Sessions = Convert-ToNumber $_.Sessions
                Actions = Convert-ToNumber $_.Actions
                MultiFeatureSessions = Convert-ToNumber $_.MultiFeatureSessions
            }
        }
)
foreach ($row in $sessionRows) {
    if ($sessionSizes -notcontains $row.Size) {
        throw "Analytics contains an unsafe session size."
    }
}
$totalSessions = [long](($sessionRows | Measure-Object Sessions -Sum).Sum)
$totalSessionActions = [long](($sessionRows | Measure-Object Actions -Sum).Sum)
$multiFeatureSessions = [long](($sessionRows | Measure-Object MultiFeatureSessions -Sum).Sum)
$sessionMedian = @($results.assistantSessionMedian)[0]
$returning = @($results.returningUsers)[0]
$newUsers = if ($null -eq $returning) { 0 } else { Convert-ToNumber $returning.NewUsers }
$returnedAfterWeek = if ($null -eq $returning) { 0 } else { Convert-ToNumber $returning.ReturnedAfterWeek }
$returnedAfterThreeWeeks = if ($null -eq $returning) { 0 } else { Convert-ToNumber $returning.ReturnedAfterThreeWeeks }
$weekdayRows = @(
    $results.weekdays |
        ForEach-Object {
            $day = [int](Convert-ToNumber $_.Day)
            if ($day -lt 0 -or $day -gt 6) {
                throw "Analytics contains an unsafe weekday."
            }
            [pscustomobject]@{
                Day = $day
                Actions = Convert-ToNumber $_.Actions
                Users = Convert-ToNumber $_.Users
            }
        }
)
$workdayRows = @($weekdayRows | Where-Object { $_.Day -ge 1 -and $_.Day -le 5 })
$weekendRows = @($weekdayRows | Where-Object { $_.Day -eq 0 -or $_.Day -eq 6 })
# Each weekday row adds up all of that weekday's actions across the whole window.
$workdayTotal = ($workdayRows | Measure-Object Actions -Sum).Sum
$weekendTotal = ($weekendRows | Measure-Object Actions -Sum).Sum
$workdayAverage = [long][Math]::Round([double]$workdayTotal / (5 * $weekdayWeeks))
$weekendAverage = [long][Math]::Round([double]$weekendTotal / (2 * $weekdayWeeks))
$habits = [ordered]@{
    windowDays = $habitDays
    minimumUsers = $habitMinimumUsers
    assistantSessions = [ordered]@{
        sessions = $totalSessions
        medianActions = if ($null -eq $sessionMedian) { 0 } else {
            [Math]::Round([double](Convert-ToNumber $sessionMedian.MedianActions), 1)
        }
        multiFeatureSessions = $multiFeatureSessions
        multiFeatureSharePct = Get-Percent $multiFeatureSessions $totalSessions
        sizes = @(
            foreach ($size in $sessionSizes) {
                $row = $sessionRows | Where-Object Size -eq $size
                $sessions = if ($null -eq $row) { 0 } else { $row.Sessions }
                $actions = if ($null -eq $row) { 0 } else { $row.Actions }
                [ordered]@{
                    size = $size
                    sessions = $sessions
                    actions = $actions
                    sessionSharePct = Get-Percent $sessions $totalSessions
                    actionSharePct = Get-Percent $actions $totalSessionActions
                }
            }
        )
    }
    featurePairs = @(
        $results.featurePairs |
            ForEach-Object {
                $first = [string]$_.First
                $second = [string]$_.Second
                if ($weights.Features -notcontains $first -or $weights.Features -notcontains $second) {
                    throw "Analytics contains an unsafe homepage-feature dimension."
                }
                [ordered]@{
                    first = $first
                    second = $second
                    sessions = Convert-ToNumber $_.Sessions
                    sharePct = Get-Percent (Convert-ToNumber $_.Sessions) $totalSessions
                }
            }
    )
    returningUsers = [ordered]@{
        newUsers = $newUsers
        returnedAfterWeek = $returnedAfterWeek
        returnedAfterWeekPct = Get-Percent $returnedAfterWeek $newUsers
        returnedAfterThreeWeeks = $returnedAfterThreeWeeks
        returnedAfterThreeWeeksPct = Get-Percent $returnedAfterThreeWeeks $newUsers
    }
    featureWait = @(
        $results.featureWait |
            ForEach-Object {
                $name = [string]$_.Feature
                if ($weights.Features -notcontains $name) {
                    throw "Analytics contains an unsafe homepage-feature dimension."
                }
                [ordered]@{
                    name = $name
                    actions = Convert-ToNumber $_.Actions
                    typicalSeconds = [Math]::Round([double](Convert-ToNumber $_.TypicalMs) / 1000, 2)
                    slowSeconds = [Math]::Round([double](Convert-ToNumber $_.SlowMs) / 1000, 1)
                }
            }
    )
    weekdays = @(
        $weekdayRows |
            Sort-Object { ($_.Day + 6) % 7 } |
            ForEach-Object {
                [ordered]@{
                    day = $weekdayNames[$_.Day]
                    actions = $_.Actions
                    users = $_.Users
                }
            }
    )
    workdayAverageActions = $workdayAverage
    weekendAverageActions = $weekendAverage
    weekdayWeeks = $weekdayWeeks
    firstAdvancedUse = @(
        $results.firstAdvancedUse |
            ForEach-Object {
                $name = [string]$_.Feature
                if ($advancedFeatures -notcontains $name) {
                    throw "Analytics contains an unsafe homepage-feature dimension."
                }
                $users = Convert-ToNumber $_.Users
                [ordered]@{
                    name = $name
                    users = $users
                    firstDay = Convert-ToNumber $_.FirstDay
                    firstWeek = Convert-ToNumber $_.FirstWeek
                    later = Convert-ToNumber $_.Later
                    firstDayPct = Get-Percent (Convert-ToNumber $_.FirstDay) $users
                }
            } |
            Sort-Object { $_.firstDayPct } -Descending
    )
}

$report = [ordered]@{
    schemaVersion = 3
    generatedAtUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    windows = [ordered]@{
        reportingDays = 90
        comparisonDays = 14
        trendWeeks = 12
        entryPointSinceUtc = $entryPointWindowStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
        entryPointMinimumUsers = $entryPointMinimumUsers
    }
    weights = $weights.Levels
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
        workUnits = $totalWorkUnits
        unweightedActions = [long](($unweightedRows | Measure-Object Actions -Sum).Sum)
    }
    comparison = [ordered]@{
        currentUsers = Convert-ToNumber $trend.Users
        previousUsers = Convert-ToNumber $trend.PreviousUsers
        userChangePct = Convert-ToNumber $trend.UserChangePct
        currentInvocations = Convert-ToNumber $trend.Invocations
        previousInvocations = Convert-ToNumber $trend.PreviousInvocations
        invocationChangePct = Convert-ToNumber $trend.InvocationChangePct
        currentWorkUnits = $currentWorkUnits
        previousWorkUnits = $previousWorkUnits
        workChangePct = Get-Percent ($currentWorkUnits - $previousWorkUnits) $previousWorkUnits
    }
    heavyWork = [ordered]@{
        users = $heavyWorkUsers
        heavyUsers = $heavyWorkHeavyUsers
        heavyUserSharePct = Get-Percent $heavyWorkHeavyUsers $heavyWorkUsers
    }
    weekly = @(
        $results.weekly |
            ForEach-Object {
                $week = ([DateTime]$_.Week).ToString("yyyy-MM-dd")
                [ordered]@{
                    week = $week
                    users = Convert-ToNumber $_.Users
                    actions = Convert-ToNumber $_.Actions
                    workUnits = [long]$workByWeek[$week]
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
                $name = [string]$_.ToolFamily
                [ordered]@{
                    name = $name
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                    sharePct = Convert-ToNumber $_.SharePct
                    workUnits = [long]$workByFamily[$name]
                    workSharePct = Get-Percent ([long]$workByFamily[$name]) $totalWorkUnits
                }
            }
    )
    heroFeatures = @(
        $results.heroFeatures |
            ForEach-Object {
                $name = [string]$_.HeroFeature
                [ordered]@{
                    name = $name
                    invocations = Convert-ToNumber $_.Invocations
                    users = Convert-ToNumber $_.Users
                    sharePct = Convert-ToNumber $_.SharePct
                    workUnits = [long]$workByFeature[$name]
                    workSharePct = Get-Percent ([long]$workByFeature[$name]) $totalWorkUnits
                }
            }
    )
    operationsByWork = @(
        $weightedRows |
            Sort-Object -Property @{ Expression = "WorkUnits"; Descending = $true }, Name |
            Select-Object -First 15 |
            ForEach-Object {
                [ordered]@{
                    name = $_.Name
                    level = $weights.LevelByName[$_.Name]
                    actions = $_.Actions
                    users = $_.Users
                    workUnits = $_.WorkUnits
                    workSharePct = Get-Percent $_.WorkUnits $totalWorkUnits
                }
            }
    )
    unweightedActions = @(
        $unweightedRows |
            Sort-Object -Property @{ Expression = "Actions"; Descending = $true }, Name |
            Select-Object -First 20 |
            ForEach-Object {
                [ordered]@{
                    name = $_.Name
                    actions = $_.Actions
                }
            }
    )
    entryPoints = @($entryPointSummaries)
    entryPointFeatures = @(
        $entryPointFeatureRows |
            Sort-Object -Property @{ Expression = { $_.entryPoint } },
                @{ Expression = { $_.workUnits }; Descending = $true },
                @{ Expression = { $_.name } }
    )
    entryPointOperations = @($entryPointOperationRows)
    habits = $habits
    downloads = $downloads

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
foreach ($feature in $report.heroFeatures) {
    if ($feature.name -notmatch '^[a-z0-9-]+$' -or $weights.Features -notcontains $feature.name) {
        throw "Analytics contains an unsafe homepage-feature dimension."
    }
}
foreach ($operation in @($report.operationsByWork) + @($report.unweightedActions) + @($report.entryPointOperations)) {
    if ($operation.name -notmatch '^[a-z0-9_/-]+$') {
        throw "Analytics contains an unsafe weighted operation dimension."
    }
}
foreach ($entryPoint in $report.entryPoints) {
    if ($entryPoints -notcontains $entryPoint.name) {
        throw "Analytics contains an unsafe entry point."
    }
}
foreach ($item in @($report.entryPointFeatures) + @($report.entryPointOperations)) {
    $summary = $report.entryPoints | Where-Object { $_.name -eq $item.entryPoint }
    if ($null -eq $summary -or -not $summary.enoughData) {
        throw "Analytics contains entry-point detail for a group below the minimum size."
    }
}
foreach ($item in $report.entryPointFeatures) {
    if ($weights.Features -notcontains $item.name) {
        throw "Analytics contains an unsafe homepage-feature dimension."
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
foreach ($channel in $report.downloads.channels) {
    if (-not $downloadChannelLabels.Contains($channel.key)) {
        throw "Analytics contains an unsafe download channel."
    }
}
foreach ($item in @($report.downloads.npmWeekly) + @($report.downloads.releases) +
    @($report.downloads.snapshots) + @($report.downloads.weeklyGains)) {
    $date = if ($item.Contains("week")) { $item.week } elseif ($item.Contains("published")) { $item.published } else { $item.date }
    if ($date -notmatch '^\d{4}-\d{2}-\d{2}$') {
        throw "Analytics contains an unsafe download date."
    }
}
$json = $report | ConvertTo-Json -Depth 8
$forbidden = @(
    '"UserId"', '"SessionId"', '"FileSessionId"', '"ClientIP"',
    '"ClientCity"', '"ClientCountryOrRegion"', '"Message"', '"StackTrace"',
    '"ExceptionType"', '"InnerExceptionTypes"', '"FailureSite"',
    '"Error"', '"ErrorMessage"', '"Stack"', '"Path"', '"Prompt"',
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

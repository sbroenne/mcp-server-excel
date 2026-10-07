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

    [string]$WeightsPath = (Join-Path $PSScriptRoot "../.github/usage-analytics-weights.json")
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "UsageAnalyticsWeights.ps1")
$weights = Read-UsageAnalyticsWeights -Path $WeightsPath
$categorizedReliabilitySinceUtc = "2026-08-28T09:25:22Z"
$categorizedReliabilityMinimumVersion = "2.0.5"
# excelcli began reporting usage, with the EntryPoint label, in release 2.0.12.
$entryPointSinceUtc = "2026-09-28T00:00:00Z"
$entryPointMinimumUsers = 10
$entryPoints = @("cli", "mcp-server")
$excludedActions = $weights.ExcludedActions
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
    reliability = @"
normalizedRequests
| where TimeGenerated >= datetime($categorizedReliabilitySinceUtc)
| extend Version=tostring(split(AppVersion, '+')[0])
| where parse_version(Version) >= parse_version('$categorizedReliabilityMinimumVersion')
| extend Outcome=tostring(Properties['Outcome']),
         FailureClass=tostring(Properties['FailureClass'])
| where Outcome in ('succeeded', 'expected-negative', 'failed')
| summarize Actions=count(),
            ExpectedNegatives=countif(Outcome == 'expected-negative'),
            Failures=countif(Outcome == 'failed'),
            InputState=countif(Outcome == 'failed' and FailureClass == 'input-state'),
            ExternalDependency=countif(Outcome == 'failed' and FailureClass == 'external-dependency'),
            TimeoutCancellation=countif(Outcome == 'failed' and FailureClass == 'timeout-cancellation'),
            ExcelRuntime=countif(Outcome == 'failed' and FailureClass == 'excel-runtime'),
            InternalProductFault=countif(Outcome == 'failed' and FailureClass == 'internal-product-fault'),
            Unclassified=countif(Outcome == 'failed' and FailureClass == 'unclassified'),
            Users=dcount(UserId) by Name
| where ExpectedNegatives > 0 or Failures > 0
| extend FailureRate=round(100.0 * Failures / Actions, 2)
| order by Failures desc, ExpectedNegatives desc
| take 20
"@
    failureClasses = @"
normalizedRequests
| where TimeGenerated >= datetime($categorizedReliabilitySinceUtc)
| extend Version=tostring(split(AppVersion, '+')[0])
| where parse_version(Version) >= parse_version('$categorizedReliabilityMinimumVersion')
| extend Outcome=tostring(Properties['Outcome']),
         FailureClass=tostring(Properties['FailureClass'])
| extend Bucket=case(
    Outcome == 'expected-negative', 'expected-negative',
    Outcome == 'failed' and FailureClass == 'input-state', 'input-state',
    Outcome == 'failed' and FailureClass == 'external-dependency', 'external-dependency',
    Outcome == 'failed' and FailureClass == 'timeout-cancellation', 'timeout-cancellation',
    Outcome == 'failed' and FailureClass == 'excel-runtime', 'excel-runtime',
    Outcome == 'failed' and FailureClass == 'internal-product-fault', 'internal-product-fault',
    Outcome == 'failed', 'unclassified',
    'succeeded')
| where Bucket != 'succeeded'
| summarize Actions=count(), Users=dcount(UserId) by Bucket
| order by Actions desc
"@
    versionReliability = @"
normalizedRequests
| where TimeGenerated >= datetime($categorizedReliabilitySinceUtc)
| extend Version=tostring(split(AppVersion, '+')[0])
| where parse_version(Version) >= parse_version('$categorizedReliabilityMinimumVersion')
| extend Outcome=tostring(Properties['Outcome']),
         FailureClass=tostring(Properties['FailureClass'])
| where Outcome in ('succeeded', 'expected-negative', 'failed')
| summarize Actions=count(),
            ExpectedNegatives=countif(Outcome == 'expected-negative'),
            Failures=countif(Outcome == 'failed'),
            InputState=countif(Outcome == 'failed' and FailureClass == 'input-state'),
            ExternalDependency=countif(Outcome == 'failed' and FailureClass == 'external-dependency'),
            TimeoutCancellation=countif(Outcome == 'failed' and FailureClass == 'timeout-cancellation'),
            ExcelRuntime=countif(Outcome == 'failed' and FailureClass == 'excel-runtime'),
            InternalProductFault=countif(Outcome == 'failed' and FailureClass == 'internal-product-fault'),
            Unclassified=countif(Outcome == 'failed' and FailureClass == 'unclassified'),
            Users=dcount(UserId) by Version
| extend FailureRate=round(100.0 * Failures / Actions, 2)
| order by Actions desc
| take 25
"@
    exceptions = @"
AppExceptions
| where TimeGenerated >= datetime($categorizedReliabilitySinceUtc)
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

$report = [ordered]@{
    schemaVersion = 3
    generatedAtUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    windows = [ordered]@{
        reportingDays = 90
        comparisonDays = 14
        trendWeeks = 12
        categorizedReliabilitySinceUtc = $categorizedReliabilitySinceUtc
        categorizedReliabilityMinimumVersion = $categorizedReliabilityMinimumVersion
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
    reliability = @(
        $results.reliability |
            Where-Object { $excludedActions -notcontains [string]$_.Name } |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Name
                    actions = Convert-ToNumber $_.Actions
                    expectedNegatives = Convert-ToNumber $_.ExpectedNegatives
                    failures = Convert-ToNumber $_.Failures
                    failureRate = Convert-ToNumber $_.FailureRate
                    inputState = Convert-ToNumber $_.InputState
                    externalDependency = Convert-ToNumber $_.ExternalDependency
                    timeoutCancellation = Convert-ToNumber $_.TimeoutCancellation
                    excelRuntime = Convert-ToNumber $_.ExcelRuntime
                    internalProductFault = Convert-ToNumber $_.InternalProductFault
                    unclassified = Convert-ToNumber $_.Unclassified
                    users = Convert-ToNumber $_.Users
                }
            }
    )
    failureClasses = @(
        $results.failureClasses |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Bucket
                    actions = Convert-ToNumber $_.Actions
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
                    expectedNegatives = Convert-ToNumber $_.ExpectedNegatives
                    failures = Convert-ToNumber $_.Failures
                    failureRate = Convert-ToNumber $_.FailureRate
                    inputState = Convert-ToNumber $_.InputState
                    externalDependency = Convert-ToNumber $_.ExternalDependency
                    timeoutCancellation = Convert-ToNumber $_.TimeoutCancellation
                    excelRuntime = Convert-ToNumber $_.ExcelRuntime
                    internalProductFault = Convert-ToNumber $_.InternalProductFault
                    unclassified = Convert-ToNumber $_.Unclassified
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
$allowedFailureClasses = @(
    "expected-negative",
    "input-state",
    "external-dependency",
    "timeout-cancellation",
    "excel-runtime",
    "internal-product-fault",
    "unclassified"
)
foreach ($failureClass in $report.failureClasses) {
    if ($allowedFailureClasses -notcontains $failureClass.name) {
        throw "Analytics contains an unsafe failure class."
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

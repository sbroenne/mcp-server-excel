<#
.SYNOPSIS
    Tests analytics aggregation, privacy boundaries, and Copilot output validation.
#>
$ErrorActionPreference = "Stop"
$updateScript = Join-Path $PSScriptRoot "Update-UsageAnalytics.ps1"
$completeScript = Join-Path $PSScriptRoot "Complete-UsageAnalyticsReport.ps1"
$interpretScript = Join-Path $PSScriptRoot "Invoke-UsageAnalyticsReport.ps1"
$persistScript = Join-Path $PSScriptRoot "Persist-UsageAnalytics.ps1"
$restoreScript = Join-Path $PSScriptRoot "Restore-UsageAnalytics.ps1"
$testRoot = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-analytics-tests-$([Guid]::NewGuid().ToString('N'))"
$utf8NoBom = [Text.UTF8Encoding]::new($false)
$testsRun = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Throws {
    param([scriptblock]$Action, [string]$ExpectedMessage)
    try {
        & $Action
    }
    catch {
        Assert-True ($_.Exception.Message -like "*$ExpectedMessage*") `
            "Expected '$ExpectedMessage', got '$($_.Exception.Message)'."
        return
    }
    throw "Expected an error containing '$ExpectedMessage'."
}

function Write-TestFile {
    param([string]$Name, [string]$Content)
    $path = Join-Path $testRoot $Name
    [IO.File]::WriteAllText($path, $Content, $utf8NoBom)
    return $path
}

New-Item -ItemType Directory -Path $testRoot | Out-Null
try {
    $fixture = @{
        overview = @(@{
            Users = 100; ToolInvocations = 1000; RepeatUserRate = 60
        })
        trend = @(@{
            Users = 50; PreviousUsers = 40; UserChangePct = 25
            Invocations = 600; PreviousInvocations = 400; InvocationChangePct = 50
        })
        weekly = @(
            @{ Week = "2026-08-09"; Users = 30; Actions = 300 },
            @{ Week = "2026-08-16"; Users = 40; Actions = 500 }
        )
        versionAdoption = @(
            @{ Week = "2026-08-09"; Version = "2.0.2"; Users = 30; SharePct = 75 },
            @{ Week = "2026-08-09"; Version = "2.0.3"; Users = 10; SharePct = 25 },
            @{ Week = "2026-08-16"; Version = "2.0.3"; Users = 40; SharePct = 100 }
        )
        operations = @(
            @{ Name = "range/get-values"; Invocations = 500; Users = 10 },
            @{ Name = "rare/action"; Invocations = 9; Users = 9 },
            @{ Name = "file/open"; Invocations = 200; Users = 20 }
        )
        families = @(
            @{ ToolFamily = "range"; Invocations = 500; Users = 10; SharePct = 50 }
        )
        heroFeatures = @(
            @{
                HeroFeature = "tables-ranges"; Invocations = 500
                Users = 10; SharePct = 50
            },
            @{
                HeroFeature = "power-query"; Invocations = 200
                Users = 5; SharePct = 20
            }
        )
        downloadSources = @{
            npm = @{
                "npm-mcp-server" = @(
                    @{ day = "2026-08-08"; downloads = 0 }
                    foreach ($offset in 0..6) {
                        @{ day = ([DateTime]"2026-08-09").AddDays($offset).ToString("yyyy-MM-dd"); downloads = 10 }
                    }
                    @{ day = "2026-08-16"; downloads = 5 }
                )
                "npm-cli" = @(
                    @{ day = "2026-08-08"; downloads = 0 }
                    foreach ($offset in 0..6) {
                        @{ day = ([DateTime]"2026-08-09").AddDays($offset).ToString("yyyy-MM-dd"); downloads = 2 }
                    }
                    @{ day = "2026-08-16"; downloads = 1 }
                )
            }
            releases = @(
                @{
                    tag = "v2.0.4"; publishedAt = "2026-08-15T10:00:00Z"; draft = $true
                    assets = @(@{ name = "ExcelMcp-MCP-Server-2.0.4-windows.zip"; downloads = 999 })
                },
                @{
                    tag = "v2.0.3"; publishedAt = "2026-08-14T10:00:00Z"; draft = $false
                    assets = @(
                        @{ name = "ExcelMcp-MCP-Server-2.0.3-windows.zip"; downloads = 40 },
                        @{ name = "excel-mcp-2.0.3.vsix"; downloads = 5 },
                        @{ name = "SHA256SUMS"; downloads = 100 },
                        @{ name = "RELEASE-INPUTS.json"; downloads = 7 }
                    )
                },
                @{
                    tag = "v2.0.2"; publishedAt = "2026-08-01T10:00:00Z"; draft = $false
                    assets = @(
                        @{ name = "ExcelMcp-MCP-Server-2.0.2-windows.zip"; downloads = 60 },
                        @{ name = "excel-mcp-2.0.2.mcpb"; downloads = 4 }
                    )
                }
            )
            vscodeInstalls = 250
        }
        actionCounts = @(
            @{ Name = "range/get-values"; Actions = 500; Users = 10 },
            @{ Name = "powerquery/refresh"; Actions = 200; Users = 5 },
            @{ Name = "rare/action"; Actions = 9; Users = 9 },
            @{ Name = "file/open"; Actions = 200; Users = 20 }
        )
        weeklyActions = @(
            @{ Week = "2026-08-09"; Name = "range/get-values"; Actions = 300 },
            @{ Week = "2026-08-16"; Name = "range/get-values"; Actions = 300 },
            @{ Week = "2026-08-16"; Name = "powerquery/refresh"; Actions = 20 },
            @{ Week = "2026-08-16"; Name = "file/open"; Actions = 50 }
        )
        comparisonActions = @(
            @{ Name = "range/get-values"; CurrentActions = 100; PreviousActions = 100 },
            @{ Name = "powerquery/refresh"; CurrentActions = 20; PreviousActions = 10 },
            @{ Name = "rare/action"; CurrentActions = 5; PreviousActions = 0 }
        )
        heavyWork = @(@{ Users = 100; HeavyUsers = 25 })
        entryPoints = @(
            @{ EntryPoint = "mcp-server"; Users = 40; Actions = 600 },
            @{ EntryPoint = "cli"; Users = 9; Actions = 50 }
        )
        entryPointActions = @(
            @{ EntryPoint = "mcp-server"; Name = "range/get-values"; Actions = 400 },
            @{ EntryPoint = "mcp-server"; Name = "powerquery/refresh"; Actions = 100 },
            @{ EntryPoint = "mcp-server"; Name = "file/open"; Actions = 100 },
            @{ EntryPoint = "cli"; Name = "range/get-values"; Actions = 50 }
        )
        assistantSessions = @(
            @{ Size = "1"; Sessions = 10; Actions = 10; MultiFeatureSessions = 0; Users = 4 },
            @{ Size = "2-10"; Sessions = 30; Actions = 150; MultiFeatureSessions = 10; Users = 20 },
            @{ Size = "201+"; Sessions = 10; Actions = 2840; MultiFeatureSessions = 10; Users = 12 }
        )
        assistantSessionMedian = @(@{ MedianActions = 5.0; Users = 30 })
        featurePairs = @(
            @{ First = "tables-ranges"; Second = "worksheets-connections"; Sessions = 20; Users = 12 }
        )
        returningUsers = @(@{ NewUsers = 200; ReturnedAfterWeek = 100; ReturnedAfterThreeWeeks = 50 })
        featureWait = @(
            @{ Feature = "power-query"; Actions = 200; Users = 12; TypicalMs = 2834.6; SlowMs = 33012.0 },
            @{ Feature = "tables-ranges"; Actions = 500; Users = 30; TypicalMs = 61.9; SlowMs = 791.3 }
        )
        weekdays = @(
            @{ Day = 0; Actions = 80; Users = 10 },
            @{ Day = 1; Actions = 400; Users = 20 },
            @{ Day = 2; Actions = 400; Users = 20 },
            @{ Day = 6; Actions = 80; Users = 5 }
        )
        firstAdvancedUse = @(
            @{ Feature = "power-query"; Users = 40; FirstDay = 30; FirstWeek = 6; Later = 4 },
            @{ Feature = "vba"; Users = 20; FirstDay = 10; FirstWeek = 5; Later = 5 }
        )
    }
    $fixturePath = Write-TestFile "fixture.json" ($fixture | ConvertTo-Json -Depth 8)
    $analyticsPath = Join-Path $testRoot "analytics.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $analyticsPath -FixturePath $fixturePath
    $analytics = Get-Content -LiteralPath $analyticsPath -Raw | ConvertFrom-Json
    Assert-True ($analytics.operations.Count -eq 2) "Low-use operations were removed."
    Assert-True ($analytics.operations[0].name -eq "range/get-values") "Expected operation was removed."
    Assert-True ($null -eq ($analytics.operations | Where-Object name -Like "file/*")) `
        "Workbook open or close actions entered the public report."
    Assert-True ($null -eq $analytics.operations[0].PSObject.Properties["successRate"]) `
        "Historical success rates entered the public report."
    Assert-True ($analytics.schemaVersion -eq 3) "Weighted analytics schema was not emitted."
    Assert-True ($null -eq $analytics.PSObject.Properties["reliability"] -and
        $null -eq $analytics.PSObject.Properties["exceptions"]) `
        "Reliability data entered the usage report; it has its own report."
    $downloads = $analytics.downloads
    Assert-True (($downloads.channels.key -join ",") -eq
        "npm-mcp-server,npm-cli,github-releases,vscode") `
        "Download channels were not published in a fixed order."
    Assert-True (($downloads.channels | Where-Object key -eq "npm-mcp-server").total -eq 75) `
        "npm downloads were not totalled from daily history."
    Assert-True (($downloads.channels | Where-Object key -eq "github-releases").total -eq 109) `
        "GitHub release downloads counted drafts, checksums, or release metadata."
    Assert-True ($downloads.npmWeekly.Count -eq 1 -and $downloads.npmWeekly[0].week -eq "2026-08-09" -and
        $downloads.npmWeekly[0].total -eq 84) `
        "npm weekly history did not keep only full weeks after the first download."
    Assert-True ($downloads.releases.Count -eq 2 -and $downloads.releases[0].version -eq "2.0.3" -and
        $downloads.releases[0].published -eq "2026-08-14" -and $downloads.releases[0].downloads -eq 45) `
        "Release downloads were not listed newest first."
    Assert-True ($downloads.snapshots.Count -eq 1 -and $downloads.weeklyGains.Count -eq 0) `
        "A first report must start the download history with one snapshot."
    Assert-True ($downloads.snapshots[0].date -eq [DateTime]::UtcNow.ToString("yyyy-MM-dd") -and
        $downloads.snapshots[0].totals.vscode -eq 250) `
        "Today's download totals were not recorded."
    Assert-True ($analytics.weekly.Count -eq 2) "Weekly usage history was not included."
    Assert-True ($analytics.versionAdoption.Count -eq 3) `
        "Weekly release adoption was not included."
    Assert-True ($analytics.versionAdoption[1].version -eq "2.0.3") `
        "Release adoption labels were not preserved."
    Assert-True ($analytics.heroFeatures[0].name -eq "tables-ranges") `
        "Homepage feature usage was not included."
    $testsRun++

    Assert-True ($analytics.weights.light -eq 1 -and $analytics.weights.medium -eq 3 -and
        $analytics.weights.heavy -eq 10) "The work levels were not published."
    Assert-True ($analytics.summary.workUnits -eq 2500) `
        "Work units did not multiply 500 light and 200 heavy actions by their levels."
    Assert-True ($analytics.summary.unweightedActions -eq 9) `
        "Actions without a weight were hidden instead of reported."
    Assert-True ($analytics.unweightedActions[0].name -eq "rare/action") `
        "The unweighted action was not named."
    $powerQuery = $analytics.heroFeatures | Where-Object name -eq "power-query"
    $tablesRanges = $analytics.heroFeatures | Where-Object name -eq "tables-ranges"
    Assert-True ($powerQuery.workSharePct -eq 80 -and $powerQuery.sharePct -eq 20) `
        "200 heavy actions out of 700 did not become 80 percent of the work."
    Assert-True ($tablesRanges.workSharePct -eq 20 -and $tablesRanges.workUnits -eq 500) `
        "Light actions did not keep their share of work."
    Assert-True ($analytics.operationsByWork[0].name -eq "powerquery/refresh" -and
        $analytics.operationsByWork[0].level -eq "heavy" -and
        $analytics.operationsByWork[0].workUnits -eq 2000) `
        "Actions were not ranked by work."
    Assert-True ($analytics.toolFamilies[0].workUnits -eq 500 -and
        $analytics.toolFamilies[0].workSharePct -eq 20) `
        "Tool families did not receive work units."
    Assert-True ($analytics.weekly[0].workUnits -eq 300 -and $analytics.weekly[1].workUnits -eq 500) `
        "Weekly work units were not calculated or included workbook open and close actions."
    Assert-True ($analytics.comparison.currentWorkUnits -eq 300 -and
        $analytics.comparison.previousWorkUnits -eq 200 -and
        $analytics.comparison.workChangePct -eq 50) `
        "The two-week work comparison is wrong."
    Assert-True ($analytics.heavyWork.heavyUserSharePct -eq 25) `
        "The share of users doing heavy work is missing."
    $testsRun++

    $mcp = $analytics.entryPoints | Where-Object name -eq "mcp-server"
    $cli = $analytics.entryPoints | Where-Object name -eq "cli"
    Assert-True ($mcp.enoughData -and $mcp.actions -eq 500 -and $mcp.workUnits -eq 1400 -and
        $mcp.actionsPerUser -eq 12.5 -and $mcp.workUnitsPerUser -eq 35) `
        "The MCP Server entry point summary is wrong."
    Assert-True (-not $cli.enoughData -and $null -eq $cli.PSObject.Properties["users"]) `
        "An entry point below the minimum group size published its numbers."
    Assert-True ($null -eq ($analytics.entryPointFeatures | Where-Object entryPoint -eq "cli") -and
        $null -eq ($analytics.entryPointOperations | Where-Object entryPoint -eq "cli")) `
        "An entry point below the minimum group size published its details."
    $mcpPowerQuery = $analytics.entryPointFeatures |
        Where-Object { $_.entryPoint -eq "mcp-server" -and $_.name -eq "power-query" }
    Assert-True ($mcpPowerQuery.actionSharePct -eq 20 -and $mcpPowerQuery.workSharePct -eq 71.43) `
        "Entry point feature shares are wrong."
    Assert-True ($analytics.windows.entryPointMinimumUsers -eq 10) `
        "The entry point minimum group size is missing."
    $testsRun++

    $habits = $analytics.habits
    Assert-True ($habits.assistantSessions.enoughData -and
        $habits.assistantSessions.sessions -eq 50 -and
        $habits.assistantSessions.medianActions -eq 5 -and
        $habits.assistantSessions.multiFeatureSharePct -eq 40) `
        "The AI assistant session summary is wrong."
    Assert-True (($habits.assistantSessions.sizes.size -join ",") -eq "1,2-10,11-50,51-200,201+") `
        "Session sizes are not listed in a fixed order with empty sizes kept."
    $singleSessions = $habits.assistantSessions.sizes | Where-Object size -eq "1"
    $emptySessions = $habits.assistantSessions.sizes | Where-Object size -eq "11-50"
    Assert-True (-not $singleSessions.enoughData -and
        $null -eq $singleSessions.PSObject.Properties["sessions"] -and
        $null -eq $singleSessions.PSObject.Properties["sessionSharePct"] -and
        $emptySessions.enoughData -and $emptySessions.sessions -eq 0) `
        "A session size used by fewer than the minimum number of users was published."
    $longSessions = $habits.assistantSessions.sizes | Where-Object size -eq "201+"
    Assert-True ($longSessions.enoughData -and
        $longSessions.sessionSharePct -eq 20 -and $longSessions.actionSharePct -eq 94.67) `
        "Long sessions did not get their share of sessions and actions."
    Assert-True ($habits.featurePairs[0].sharePct -eq 40 -and
        $null -eq $habits.featurePairs[0].PSObject.Properties["users"]) `
        "Areas used together are wrong or publish a user count."
    Assert-True ($habits.returningUsers.enoughData -and
        $habits.returningUsers.returnedAfterWeekPct -eq 50 -and
        $habits.returningUsers.returnedAfterThreeWeeksPct -eq 25) `
        "Returning user shares are wrong."
    Assert-True ($habits.featureWait[0].name -eq "power-query" -and
        $habits.featureWait[0].typicalSeconds -eq 2.83 -and
        $habits.featureWait[0].slowSeconds -eq 33) `
        "Typical waits were not converted to seconds."
    $saturday = $habits.weekdays | Where-Object day -eq "Saturday"
    Assert-True (($habits.weekdays.day -join ",") -eq "Monday,Tuesday,Saturday,Sunday" -and
        $habits.workdayAverageActions -eq 20 -and $habits.weekendAverageActions -eq 10) `
        "Weekday use is not in Monday-first order or the per-day averages are wrong."
    Assert-True (-not $saturday.enoughData -and
        $null -eq $saturday.PSObject.Properties["actions"] -and
        $null -eq $saturday.PSObject.Properties["users"]) `
        "A weekday used by fewer than the minimum number of users was published."
    Assert-True ($habits.firstAdvancedUse[0].name -eq "power-query" -and
        $habits.firstAdvancedUse[0].firstDayPct -eq 75 -and
        $habits.firstAdvancedUseWindowDays -eq 60) `
        "First use of advanced areas is wrong or does not state its window."
    $habitQuerySource = [IO.File]::ReadAllText($updateScript)
    $firstUseQuery = [regex]::Match($habitQuerySource, '(?s)firstAdvancedUse = @".*?"@').Value
    Assert-True ($firstUseQuery -match 'FirstSeen > ago\(\$\{firstAdvancedCohortDays\}d\)') `
        "First use of advanced areas does not limit itself to people who started inside the window."
    $sessionQuery = [regex]::Match($habitQuerySource, '(?s)assistantSessions = @".*?"@').Value
    Assert-True ($sessionQuery -match 'Users=dcount\(UserId\)') `
        "Session sizes do not count users, so small groups cannot be hidden."
    foreach ($queryName in "assistantSessions", "assistantSessionMedian", "featurePairs") {
        $query = [regex]::Match($habitQuerySource, "(?s)$queryName = @`".*?`"@").Value
        Assert-True ($query -match 'by UserId, SessionId\r?\n') `
            "$queryName must group by person and session; SessionId alone repeats across people."
    }

    $smallHabitFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $smallHabitFixture.assistantSessionMedian[0].Users = 9
    $smallHabitFixture.returningUsers[0].NewUsers = 9
    $smallHabitFixture.returningUsers[0].ReturnedAfterWeek = 3
    $smallHabitFixture.returningUsers[0].ReturnedAfterThreeWeeks = 1
    $smallHabitPath = Write-TestFile "small-habit-fixture.json" ($smallHabitFixture | ConvertTo-Json -Depth 8)
    $smallHabitOutput = Join-Path $testRoot "small-habits.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $smallHabitOutput -FixturePath $smallHabitPath
    $smallHabits = (Get-Content -LiteralPath $smallHabitOutput -Raw | ConvertFrom-Json).habits
    Assert-True (-not $smallHabits.assistantSessions.enoughData -and
        $null -eq $smallHabits.assistantSessions.PSObject.Properties["sessions"] -and
        $null -eq $smallHabits.assistantSessions.PSObject.Properties["sizes"]) `
        "Session figures from fewer than the minimum number of users were published."
    Assert-True (-not $smallHabits.returningUsers.enoughData -and
        $null -eq $smallHabits.returningUsers.PSObject.Properties["newUsers"] -and
        $null -eq $smallHabits.returningUsers.PSObject.Properties["returnedAfterWeekPct"]) `
        "A returning-user group smaller than the minimum was published."
    $yearFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $yearDays = @(
        for ($offset = 0; $offset -lt 60 * 7; $offset++) {
            @{ day = ([DateTime]"2025-08-10").AddDays($offset).ToString("yyyy-MM-dd"); downloads = 1 }
        }
    )
    $yearFixture.downloadSources.npm."npm-mcp-server" = $yearDays
    $yearFixture.downloadSources.npm."npm-cli" = $yearDays
    $yearPath = Write-TestFile "year-fixture.json" ($yearFixture | ConvertTo-Json -Depth 8)
    $yearOutput = Join-Path $testRoot "year.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $yearOutput -FixturePath $yearPath
    $yearWeekly = (Get-Content -LiteralPath $yearOutput -Raw | ConvertFrom-Json).downloads.npmWeekly
    Assert-True ($yearWeekly.Count -eq 52 -and $yearWeekly[-1].total -eq 14) `
        "npm weekly history does not cover the last 12 months."
    $unsafeHabitCases = @{
        "session size" = { param($f) $f.assistantSessions[0].Size = "huge" }
        "homepage-feature" = { param($f) $f.featurePairs[0].Second = "private-workbook" }
        "weekday" = { param($f) $f.weekdays[0].Day = 9 }
    }
    foreach ($case in $unsafeHabitCases.GetEnumerator()) {
        $unsafeHabitFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
        & $case.Value $unsafeHabitFixture
        $unsafeHabitPath = Write-TestFile "unsafe-habit-fixture.json" ($unsafeHabitFixture | ConvertTo-Json -Depth 8)
        Assert-Throws -ExpectedMessage $case.Key -Action {
            & $updateScript -WorkspaceId "fixture" `
                -OutputPath (Join-Path $testRoot "unsafe-habit-report.json") `
                -FixturePath $unsafeHabitPath
        }
    }
    $testsRun++

    $largeCliFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $largeCliFixture.entryPoints[1].Users = 10
    $largeCliPath = Write-TestFile "large-cli-fixture.json" ($largeCliFixture | ConvertTo-Json -Depth 8)
    $largeCliAnalyticsPath = Join-Path $testRoot "large-cli.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $largeCliAnalyticsPath -FixturePath $largeCliPath
    $largeCli = (Get-Content -LiteralPath $largeCliAnalyticsPath -Raw | ConvertFrom-Json).entryPoints |
        Where-Object name -eq "cli"
    Assert-True ($largeCli.enoughData -and $largeCli.users -eq 10 -and $largeCli.workUnits -eq 50) `
        "An entry point at the minimum group size was hidden."
    $testsRun++

    . (Join-Path $PSScriptRoot "UsageAnalyticsWeights.ps1")
    $weights = Read-UsageAnalyticsWeights -Path (Join-Path $PSScriptRoot "../.github/usage-analytics-weights.json")
    Assert-True ($weights.ToolMap["rangeformat"] -eq "range_format") `
        "CLI category names are not mapped to MCP tool names."
    Assert-True ($weights.ToolMap["range_read"] -eq "range") `
        "Read-only MCP tool names are not mapped to their base tool."
    Assert-True ($weights.SplitMap["sheet/set-tab-color"] -eq "worksheet_style/set-tab-color" -and
        $weights.SplitMap["sheet/list"] -eq "worksheet/list") `
        "CLI sheet actions are not split between worksheet tools."
    Assert-True ($weights.ToolMap["session"] -eq "file" -and $weights.ExcludedActions -contains "file/open") `
        "CLI session open and close are not treated like workbook open and close."
    Assert-True ($weights.HeavyNames -contains "powerquery/refresh" -and
        $weights.HeavyNames -notcontains "range/get-values") `
        "Heavy actions were not identified."
    $prelude = New-UsageAnalyticsQueryPrelude -Weights $weights
    Assert-True ($prelude.Contains("'rangeformat', 'range_format', 'tables-ranges'")) `
        "The query lookup does not combine CLI and MCP spellings."
    Assert-True ($prelude.Contains("'sheet/set-tab-color', 'worksheet_style/set-tab-color', 'worksheets-connections'")) `
        "The query lookup does not split CLI sheet actions."
    Assert-True ($prelude.Contains("coalesce(tostring(Properties['EntryPoint']), 'mcp-server')")) `
        "Rows recorded before the entry point label are not counted as MCP Server."
    Assert-True ($prelude -match "let excludedNames = dynamic\(\['file/open', 'file/close'\]\)") `
        "Excluded actions are not removed after names are combined."
    $testsRun++

    $validWeights = Get-Content -LiteralPath (Join-Path $PSScriptRoot "../.github/usage-analytics-weights.json") -Raw
    $invalidWeights = @{
        "unknown level" = $validWeights.Replace('"get-values": "light"', '"get-values": "enormous"')
        "positive whole number" = $validWeights.Replace('"heavy": 10', '"heavy": 0')
        "repeat" = $validWeights.Replace('"get-values": "light",', '"get-values": "light", "get-values": "heavy",')
        "unknown tool" = $validWeights.Replace('"rangeformat": ["range_format"]', '"rangeformat": ["range_formats"]')
        "unknown feature" = $validWeights.Replace('"feature": "power-query"', '"feature": "power-queries"')
    }
    foreach ($case in $invalidWeights.GetEnumerator()) {
        Assert-True ($case.Value -ne $validWeights) "Invalid weights case '$($case.Key)' did not change the file."
        $invalidWeightsPath = Write-TestFile "invalid-weights.json" $case.Value
        Assert-Throws -ExpectedMessage $case.Key -Action {
            & $updateScript -WorkspaceId "fixture" `
                -OutputPath (Join-Path $testRoot "invalid-weights-report.json") `
                -FixturePath $fixturePath `
                -WeightsPath $invalidWeightsPath
        }
    }
    $testsRun++

    $querySource = [IO.File]::ReadAllText($updateScript)
    Assert-True ($querySource -match 'iif\(\s*count\(\) == 0,\s*0\.0,') `
        "The repeat-use query does not guard an empty reporting window."
    Assert-True ($querySource -match 'iif\(\s*PreviousUsers == 0,\s*0\.0,') `
        "The user comparison does not guard an empty previous window."
    Assert-True ($querySource -match 'iif\(\s*PreviousInvocations == 0,\s*0\.0,') `
        "The action comparison does not guard an empty previous window."
    $testsRun++

    $today = [DateTime]::UtcNow.ToString("yyyy-MM-dd")
    $previousReport = $analytics | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    $previousReport.downloads.snapshots = @(
        [pscustomobject]@{
            date = "2026-08-01"
            totals = [pscustomobject]@{
                "nuget-mcp-server" = 400; "github-releases" = 100; vscode = 200
            }
        },
        [pscustomobject]@{
            date = $today
            totals = [pscustomobject]@{
                "github-releases" = 999; vscode = 999
            }
        }
    )
    $previousReportPath = Write-TestFile "previous-report.json" ($previousReport | ConvertTo-Json -Depth 10)
    $carriedPath = Join-Path $testRoot "carried.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $carriedPath -FixturePath $fixturePath `
        -PreviousReportPath $previousReportPath
    $carried = (Get-Content -LiteralPath $carriedPath -Raw | ConvertFrom-Json).downloads
    Assert-True ($carried.snapshots.Count -eq 2 -and $carried.snapshots[0].date -eq "2026-08-01") `
        "Earlier download snapshots were not carried forward."
    Assert-True ($carried.snapshots[1].date -eq $today -and $carried.snapshots[1].totals.vscode -eq 250) `
        "A second run on the same day did not replace that day's snapshot."
    Assert-True ($carried.weeklyGains.Count -eq 1 -and $carried.weeklyGains[0].week -eq "2026-08-01" -and
        $carried.weeklyGains[0].total -eq 59 -and $carried.weeklyGains[0].channels.vscode -eq 50 -and
        $null -eq $carried.weeklyGains[0].channels.PSObject.Properties["nuget-mcp-server"] -and
        $null -eq $carried.snapshots[0].totals.PSObject.Properties["nuget-mcp-server"]) `
        "Download gains were not calculated between snapshots, or a retired channel was kept."
    $testsRun++

    $bootstrapCarriedPath = Join-Path $testRoot "bootstrap-carried.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $bootstrapCarriedPath -FixturePath $fixturePath `
        -PreviousReportPath (Join-Path $PSScriptRoot "../.github/usage-analytics.json")
    $bootstrapCarried = (Get-Content -LiteralPath $bootstrapCarriedPath -Raw | ConvertFrom-Json).downloads
    Assert-True ($bootstrapCarried.snapshots.Count -ge 1 -and
        $bootstrapCarried.snapshots[-1].date -eq $today) `
        "The checked-in report could not seed the download history."
    $testsRun++

    $brokenPreviousReport = $previousReport | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    $brokenPreviousReport.downloads.snapshots[0].totals.PSObject.Properties.Remove("vscode")
    $brokenPreviousPath = Write-TestFile "broken-previous.json" ($brokenPreviousReport | ConvertTo-Json -Depth 10)
    Assert-Throws -ExpectedMessage "snapshot is missing 'vscode'" -Action {
        & $updateScript -WorkspaceId "fixture" -OutputPath (Join-Path $testRoot "broken.json") `
            -FixturePath $fixturePath -PreviousReportPath $brokenPreviousPath
    }
    Assert-Throws -ExpectedMessage "does not exist" -Action {
        & $updateScript -WorkspaceId "fixture" -OutputPath (Join-Path $testRoot "missing.json") `
            -FixturePath $fixturePath -PreviousReportPath (Join-Path $testRoot "no-such-report.json")
    }
    $testsRun++

    $unsafeDownloadFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $unsafeDownloadFixture.downloadSources.releases[1].tag = "v2.0.3 see private notes"
    $unsafeDownloadPath = Write-TestFile "unsafe-download-fixture.json" `
        ($unsafeDownloadFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe release version" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "unsafe-download.json") `
            -FixturePath $unsafeDownloadPath
    }
    $testsRun++

    $invalidFeatureFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $invalidFeatureFixture.families = @()
    $invalidFeatureFixture.heroFeatures[0].HeroFeature = "unsafe feature"
    $invalidFeaturePath = Write-TestFile "invalid-feature-fixture.json" `
        ($invalidFeatureFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe homepage-feature dimension" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "invalid-feature.json") `
            -FixturePath $invalidFeaturePath
    }

    $invalidReleaseFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $invalidReleaseFixture.weekly = @()
    $invalidReleaseFixture.versionAdoption[0].Version = "unsafe version"
    $invalidReleasePath = Write-TestFile "invalid-release-fixture.json" `
        ($invalidReleaseFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe release-adoption dimension" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "invalid-release.json") `
            -FixturePath $invalidReleasePath
    }
    $testsRun++

    $nonNumericFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $nonNumericFixture.overview[0].Users = $true
    $nonNumericPath = Write-TestFile "non-numeric-fixture.json" `
        ($nonNumericFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "non-numeric value" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "non-numeric.json") `
            -FixturePath $nonNumericPath
    }
    $testsRun++

    $interpretation = @"
## What changed

Users increased by 25 percent while the report covered 1,000 actions.

## How people use it

Release 2.0.3 was used by 40 people in the latest full week.

## Where people get it

The VS Code extension shows 250 installs, and npm recorded 84 downloads in one full week. Downloads count downloads, not people.

## What we will improve

Review the 200 heavy data actions before changing behavior.
"@
    $interpretationPath = Write-TestFile "interpretation.md" $interpretation
    $reportPath = Join-Path $testRoot "report.json"
    & $completeScript `
        -AnalyticsPath $analyticsPath `
        -InterpretationPath $interpretationPath `
        -OutputPath $reportPath
    $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
    Assert-True ($report.interpretation -like "*What changed*") "Interpretation was not added."
    Assert-True ($report.interpretationModel -eq "GitHub Copilot CLI") "Model label is missing."
    $testsRun++

    $retryInterpretationPath = Join-Path $testRoot "retry-interpretation.md"
    $retryReportPath = Join-Path $testRoot "retry-report.json"
    $copilotRequests = [Collections.Generic.List[string]]::new()
    $copilotInvoker = {
        param([string[]]$Arguments)
        $copilotRequests.Add(($Arguments -join " "))
        $content = if ($copilotRequests.Count -eq 1) {
            $interpretation + ("x" * 4000)
        }
        else {
            $interpretation
        }
        [IO.File]::WriteAllText($retryInterpretationPath, $content, $utf8NoBom)
        [pscustomobject]@{ ExitCode = 0; Output = @() }
    }
    & $interpretScript `
        -AnalyticsPath $analyticsPath `
        -InterpretationPath $retryInterpretationPath `
        -OutputPath $retryReportPath `
        -CopilotInvoker $copilotInvoker
    $retryReport = Get-Content -LiteralPath $retryReportPath -Raw | ConvertFrom-Json
    Assert-True ($copilotRequests.Count -eq 2) `
        "An oversized interpretation was not regenerated exactly once."
    Assert-True ($copilotRequests[0] -like "*between 100 and 3,500 characters*") `
        "The initial prompt does not leave room below the validation limit."
    Assert-True ($copilotRequests[1] -like "*failed validation*") `
        "The retry prompt does not explain why another draft is required."
    Assert-True (@($copilotRequests | Where-Object { $_ -like "*--model claude-opus-5.5 *" }).Count -eq 2) `
        "Every interpretation request must pin the report model."
    Assert-True ($retryReport.interpretation -eq $interpretation.Trim()) `
        "The regenerated interpretation was not assembled into the report."
    $testsRun++

    $failedInterpretationPath = Join-Path $testRoot "failed-interpretation.md"
    $failedReportPath = Join-Path $testRoot "failed-report.json"
    $failedRequests = [Collections.Generic.List[string]]::new()
    $invalidCopilotInvoker = {
        param([string[]]$Arguments)
        $failedRequests.Add(($Arguments -join " "))
        [IO.File]::WriteAllText(
            $failedInterpretationPath,
            $interpretation + ("x" * 4000),
            $utf8NoBom)
        [pscustomobject]@{ ExitCode = 0; Output = @() }
    }
    Assert-Throws -ExpectedMessage "failed validation after 2 attempts" -Action {
        & $interpretScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $failedInterpretationPath `
            -OutputPath $failedReportPath `
            -MaxAttempts 2 `
            -CopilotInvoker $invalidCopilotInvoker
    }
    Assert-True ($failedRequests.Count -eq 2) `
        "Invalid interpretation generation did not stop at the configured attempt limit."
    Assert-True (-not (Test-Path -LiteralPath $failedReportPath)) `
        "An invalid interpretation produced a report."
    $testsRun++

    $requests = [Collections.Generic.List[string]]::new()
    $persistInvoker = {
        param([string[]]$Arguments)
        $request = $Arguments -join " "
        $requests.Add($request)
        if ($request -eq "api repos/owner/repository/git/ref/heads/analytics-data" -or
            $request -eq "api repos/owner/repository/contents/.github/usage-analytics.json?ref=analytics-data") {
            return [pscustomobject]@{ ExitCode = 1; Output = @("gh: Not Found (HTTP 404)") }
        }
        return [pscustomobject]@{ ExitCode = 0; Output = @() }
    }
    & $persistScript `
        -Repository "owner/repository" `
        -Branch "analytics-data" `
        -ReportPath $reportPath `
        -RemotePath ".github/usage-analytics.json" `
        -CommitSha "0123456789abcdef" `
        -TempPath $testRoot `
        -ApiInvoker $persistInvoker
    Assert-True (($requests | Where-Object {
        $_ -eq "api repos/owner/repository/contents/.github/usage-analytics.json?ref=analytics-data"
    }).Count -eq 1) `
        "Persist script split the branch query into a second API endpoint."
    Assert-True (($requests | Where-Object {
        $_ -like "api --silent --method PUT repos/owner/repository/contents/.github/usage-analytics.json --input *"
    }).Count -eq 1) `
        "Persist script did not write through the Contents API."
    Assert-True (($requests | Where-Object { $_ -like "*refs/heads/analytics-data*" }).Count -eq 1) `
        "Persist script did not create the dedicated data branch."
    $testsRun++

    $bootstrapPath = Write-TestFile "bootstrap.json" ($report | ConvertTo-Json -Depth 10)
    $restoredText = [IO.File]::ReadAllText($reportPath)
    $restoredContent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($restoredText))
    $restoreRequests = [Collections.Generic.List[string]]::new()
    $restoreInvoker = {
        param([string[]]$Arguments)
        $restoreRequests.Add(($Arguments -join " "))
        [pscustomobject]@{
            ExitCode = 0
            Output = @((@{ content = $restoredContent } | ConvertTo-Json -Compress))
        }
    }
    & $restoreScript `
        -Repository "owner/repository" `
        -Branch "analytics-data" `
        -ReportPath $bootstrapPath `
        -RemotePath ".github/usage-analytics.json" `
        -ApiInvoker $restoreInvoker
    Assert-True ($restoreRequests[0] -eq
        "api repos/owner/repository/contents/.github/usage-analytics.json?ref=analytics-data") `
        "Restore script split the branch query into a second API endpoint."
    Assert-True ([IO.File]::ReadAllText($bootstrapPath) -eq $restoredText) `
        "Restore script did not replace the bootstrap report."
    $testsRun++

    $schemaTwoReport = $report | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    $schemaTwoReport.schemaVersion = 2
    $schemaTwoText = $schemaTwoReport | ConvertTo-Json -Depth 10
    $restoredContent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($schemaTwoText))
    & $restoreScript `
        -Repository "owner/repository" `
        -Branch "analytics-data" `
        -ReportPath $bootstrapPath `
        -RemotePath ".github/usage-analytics.json" `
        -ApiInvoker $restoreInvoker
    Assert-True ([IO.File]::ReadAllText($bootstrapPath) -eq $schemaTwoText) `
        "Restore script rejected a report from before work weighting."
    $testsRun++

    $unsupported = $interpretation.Replace("84 downloads", "999 downloads")
    $unsupportedPath = Write-TestFile "unsupported.md" $unsupported
    Assert-Throws -ExpectedMessage "unsupported numeric claim '999'" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $unsupportedPath `
            -OutputPath (Join-Path $testRoot "unsupported.json")
    }
    $testsRun++

    $unsupportedVersion = $interpretation.Replace("Release 2.0.3", "Release 2.0.9")
    $unsupportedVersionPath = Write-TestFile "unsupported-version.md" $unsupportedVersion
    Assert-Throws -ExpectedMessage "unsupported numeric claim '2.0.9'" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $unsupportedVersionPath `
            -OutputPath (Join-Path $testRoot "unsupported-version.json")
    }
    $testsRun++

    $jargonInterpretation = $interpretation.Replace(
        "heavy data actions",
        "sanitized AggregateException records")
    $jargonInterpretationPath = Write-TestFile "jargon-interpretation.md" $jargonInterpretation
    Assert-Throws -ExpectedMessage "forbidden technical jargon" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $jargonInterpretationPath `
            -OutputPath (Join-Path $testRoot "jargon-report.json")
    }
    $testsRun++

    $unsafeInterpretation = $interpretation.Replace(
        "before changing behavior.",
        "before changing behavior for customer@example.com.")
    $unsafeInterpretationPath = Write-TestFile "unsafe-interpretation.md" $unsafeInterpretation
    Assert-Throws -ExpectedMessage "forbidden email content" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $unsafeInterpretationPath `
            -OutputPath (Join-Path $testRoot "unsafe-report.json")
    }
    $testsRun++
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Usage analytics tests passed ($testsRun checks)."

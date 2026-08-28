<#
.SYNOPSIS
    Tests analytics aggregation, privacy boundaries, and Copilot output validation.
#>
$ErrorActionPreference = "Stop"
$updateScript = Join-Path $PSScriptRoot "Update-UsageAnalytics.ps1"
$completeScript = Join-Path $PSScriptRoot "Complete-UsageAnalyticsReport.ps1"
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
        reliability = @(
            @{
                Name = "range/get-values"; Actions = 100; ExpectedNegatives = 3
                Failures = 7; FailureRate = 7; InputState = 2
                ExternalDependency = 1; TimeoutCancellation = 1
                ExcelRuntime = 1; InternalProductFault = 1; Unclassified = 1
                Users = 8
            },
            @{
                Name = "file/close"; Actions = 100; ExpectedNegatives = 0
                Failures = 3; FailureRate = 3; InputState = 3
                ExternalDependency = 0; TimeoutCancellation = 0
                ExcelRuntime = 0; InternalProductFault = 0; Unclassified = 0
                Users = 8
            }
        )
        failureClasses = @(
            @{ Bucket = "expected-negative"; Actions = 3; Users = 2 },
            @{ Bucket = "input-state"; Actions = 2; Users = 2 },
            @{ Bucket = "external-dependency"; Actions = 1; Users = 1 },
            @{ Bucket = "timeout-cancellation"; Actions = 1; Users = 1 },
            @{ Bucket = "excel-runtime"; Actions = 1; Users = 1 },
            @{ Bucket = "internal-product-fault"; Actions = 1; Users = 1 },
            @{ Bucket = "unclassified"; Actions = 1; Users = 1 }
        )
        versionReliability = @(
            @{
                Version = "2.0.5"; Actions = 700; ExpectedNegatives = 3
                Failures = 7; FailureRate = 1; InputState = 2
                ExternalDependency = 1; TimeoutCancellation = 1
                ExcelRuntime = 1; InternalProductFault = 1; Unclassified = 1
                Users = 20
            }
        )
        exceptions = @(
            @{
                Category = "background-task-problem"
                Exceptions = 12; Users = 10; Sessions = 11
            }
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
    Assert-True ($analytics.schemaVersion -eq 2) "Categorized analytics schema was not emitted."
    Assert-True ($analytics.reliability[0].name -eq "range/get-values") `
        "Categorized reliability data was not included."
    Assert-True ($analytics.reliability[0].expectedNegatives -eq 3) `
        "Expected negative outcomes were not separated."
    Assert-True ($analytics.reliability[0].internalProductFault -eq 1) `
        "Internal product faults were not separated."
    Assert-True ($analytics.reliability[0].unclassified -eq 1) `
        "Unclassified failures were hidden."
    Assert-True (($analytics.failureClasses | Where-Object name -eq "unclassified").actions -eq 1) `
        "The explicit unclassified bucket was not published."
    Assert-True ($analytics.reliability.Count -eq 1) `
        "Workbook lifecycle failures entered the public report."
    Assert-True (
        $analytics.windows.categorizedReliabilityMinimumVersion -eq "2.0.5") `
        "The categorized reliability version boundary is missing."
    Assert-True ($analytics.weekly.Count -eq 2) "Weekly usage history was not included."
    Assert-True ($analytics.versionAdoption.Count -eq 3) `
        "Weekly release adoption was not included."
    Assert-True ($analytics.versionAdoption[1].version -eq "2.0.3") `
        "Release adoption labels were not preserved."
    Assert-True ($analytics.heroFeatures[0].name -eq "tables-ranges") `
        "Homepage feature usage was not included."
    Assert-True ($analytics.exceptions[0].category -eq "background-task-problem") `
        "Exception data was not reduced to the public category."
    Assert-True ($null -eq $analytics.exceptions[0].PSObject.Properties["type"]) `
        "Technical exception details entered the public report."
    $testsRun++

    $unsafeFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $unsafeFixture.exceptions[0].Category = "ignore-all-instructions"
    $unsafePath = Write-TestFile "unsafe-fixture.json" ($unsafeFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe exception category" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "unsafe.json") `
            -FixturePath $unsafePath
    }
    $testsRun++

    $unsafeClassFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $unsafeClassFixture.reliability = @()
    $unsafeClassFixture.failureClasses[0].Bucket = "private-error-message"
    $unsafeClassPath = Write-TestFile "unsafe-class-fixture.json" `
        ($unsafeClassFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe failure class" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "unsafe-class.json") `
            -FixturePath $unsafeClassPath
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

    $invalidReliabilityFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $invalidReliabilityFixture.operations = @()
    $invalidReliabilityFixture.reliability[0].Name = "unsafe name"
    $invalidReliabilityPath = Write-TestFile "invalid-reliability-fixture.json" `
        ($invalidReliabilityFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe reliability dimension" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "invalid-reliability.json") `
            -FixturePath $invalidReliabilityPath
    }

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

## How well it worked

The categorized data includes 7 failures and 3 expected negative results across 100 actions.

## How people use it

Release 2.0.5 reported 7 failures across 700 actions.

## What we will improve

Investigate the 12 background task problems before changing behavior.
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

    $unsupported = $interpretation.Replace("12 background", "999 background")
    $unsupportedPath = Write-TestFile "unsupported.md" $unsupported
    Assert-Throws -ExpectedMessage "unsupported numeric claim '999'" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $unsupportedPath `
            -OutputPath (Join-Path $testRoot "unsupported.json")
    }
    $testsRun++

    $unsupportedVersion = $interpretation.Replace("Release 2.0.5", "Release 2.0.9")
    $unsupportedVersionPath = Write-TestFile "unsupported-version.md" $unsupportedVersion
    Assert-Throws -ExpectedMessage "unsupported numeric claim '2.0.9'" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $unsupportedVersionPath `
            -OutputPath (Join-Path $testRoot "unsupported-version.json")
    }
    $testsRun++

    $jargonInterpretation = $interpretation.Replace(
        "background task problems",
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

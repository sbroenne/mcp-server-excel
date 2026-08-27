<#
.SYNOPSIS
    Tests analytics aggregation, cohort privacy, and Copilot output validation.
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
            Users = 100; Sessions = 200; ActivatedSessions = 150
            ActivationRate = 75; ToolInvocations = 1000
            RepeatUserRate = 60; MultiSessionRate = 70
        })
        trend = @(@{
            Users = 50; PreviousUsers = 40; UserChangePct = 25
            Sessions = 80; PreviousSessions = 60; SessionChangePct = 33.33
            Invocations = 600; PreviousInvocations = 400; InvocationChangePct = 50
        })
        operations = @(
            @{ Name = "range/get-values"; Invocations = 500; Users = 10; SuccessRate = 99.8; P50Ms = 10; P95Ms = 50; P99Ms = 100 },
            @{ Name = "rare/action"; Invocations = 9; Users = 9; SuccessRate = 100; P50Ms = 1; P95Ms = 2; P99Ms = 3 }
        )
        families = @(
            @{ ToolFamily = "range"; Invocations = 500; Users = 10; SharePct = 50; SuccessRate = 99.8; P95Ms = 50 }
        )
        versions = @(
            @{
                Version = "2.0.2"; Invocations = 700; Sessions = 80
                ActivatedSessions = 60; ActivationRate = 75; Users = 20
            }
        )
        exceptions = @(
            @{
                Source = "TaskScheduler.UnobservedTaskException"
                ExceptionType = "AggregateException"
                InnerExceptionTypes = "COMException"
                FailureSite = "Sbroenne.ExcelMcp.Core.Commands.VbaCommands.Run"
                Exceptions = 12; Users = 10; Sessions = 11
            }
        )
    }
    $fixturePath = Write-TestFile "fixture.json" ($fixture | ConvertTo-Json -Depth 8)
    $analyticsPath = Join-Path $testRoot "analytics.json"
    & $updateScript -WorkspaceId "fixture" -OutputPath $analyticsPath -FixturePath $fixturePath
    $analytics = Get-Content -LiteralPath $analyticsPath -Raw | ConvertFrom-Json
    Assert-True ($analytics.operations.Count -eq 1) "Small cohorts were not suppressed."
    Assert-True ($analytics.operations[0].name -eq "range/get-values") "Expected operation was removed."
    Assert-True ($analytics.privacy.minimumUsersPerDimension -eq 10) "Privacy threshold was not recorded."
    $testsRun++

    $unsafeFixture = $fixture | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $unsafeFixture.exceptions[0].FailureSite = "C:\Users\customer\secret.xlsx"
    $unsafePath = Write-TestFile "unsafe-fixture.json" ($unsafeFixture | ConvertTo-Json -Depth 8)
    Assert-Throws -ExpectedMessage "unsafe exception dimension" -Action {
        & $updateScript -WorkspaceId "fixture" `
            -OutputPath (Join-Path $testRoot "unsafe.json") `
            -FixturePath $unsafePath
    }
    $testsRun++

    $interpretation = @"
## What changed

Users increased by 25 percent while invocations increased by 50 percent.

## Reliability and performance

The leading operation had a 99.8 percent success rate and a 50 ms tail.

## Adoption

Version 2.0.2 served 20 users.

## Recommendations

Investigate the 12 sanitized AggregateException records before changing behavior.
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
        $_ -like "api --silent --method PUT repos/owner/repository/contents/.github/usage-analytics.json --input *"
    }).Count -eq 1) `
        "Persist script did not write through the Contents API."
    Assert-True (($requests | Where-Object { $_ -like "*refs/heads/analytics-data*" }).Count -eq 1) `
        "Persist script did not create the dedicated data branch."
    $testsRun++

    $bootstrapPath = Write-TestFile "bootstrap.json" ($report | ConvertTo-Json -Depth 10)
    $restoredText = [IO.File]::ReadAllText($reportPath)
    $restoredContent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($restoredText))
    $restoreInvoker = {
        param([string[]]$Arguments)
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
    Assert-True ([IO.File]::ReadAllText($bootstrapPath) -eq $restoredText) `
        "Restore script did not replace the bootstrap report."
    $testsRun++

    $unsupported = $interpretation.Replace("12 sanitized", "999 sanitized")
    $unsupportedPath = Write-TestFile "unsupported.md" $unsupported
    Assert-Throws -ExpectedMessage "unsupported numeric claim '999'" -Action {
        & $completeScript `
            -AnalyticsPath $analyticsPath `
            -InterpretationPath $unsupportedPath `
            -OutputPath (Join-Path $testRoot "unsupported.json")
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

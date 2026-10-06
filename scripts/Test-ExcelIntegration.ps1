<#
.SYNOPSIS
Runs affected real-Excel integration tests and retains reconciled evidence.
.DESCRIPTION
Use focused mode for tests affected by a change. Full mode is an explicit
complete-suite check, not a routine development step. OnDemand diagnostics
are separate from ordinary integration tests.
.EXAMPLE
.\scripts\Test-ExcelIntegration.ps1 -Project Service -Filter 'Feature=Tables&RunType!=OnDemand'
.EXAMPLE
.\scripts\Test-ExcelIntegration.ps1 -Full
.EXAMPLE
.\scripts\Test-ExcelIntegration.ps1 -Full -ContinueOnFailure
#>
[CmdletBinding(DefaultParameterSetName = 'Focused')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Full')]
    [switch]$Full,
    [Parameter(ParameterSetName = 'Full')]
    [switch]$ContinueOnFailure,
    [Parameter(Mandatory, ParameterSetName = 'Focused')]
    [ValidateSet('Core', 'Service', 'CLI', 'McpServer', 'ComInterop', 'SkillGeneration', 'ScriptSafety', 'Packaging', 'Diagnostics')]
    [string]$Project,
    [Parameter(Mandatory, ParameterSetName = 'Focused')]
    [ValidateNotNullOrEmpty()]
    [string]$Filter,
    [switch]$DiscoverOnly,
    [string]$ResultsDirectory,
    [ValidateRange(1, 86400)]
    [int]$StageTimeoutSeconds = 7200,
    [ValidateRange(1, 3600)]
    [int]$HangTimeoutSeconds = 600
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

function Invoke-ExcelBehaviorProcess {
    param(
        [string]$Executable = 'dotnet',
        [string[]]$Arguments,
        [string]$WorkingDirectory,
        [string]$LogBase,
        [int]$DeadlineSeconds
    )
    @{ executable = $Executable; arguments = $Arguments; deadlineSeconds = $DeadlineSeconds } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath "$LogBase.command.json"
    $info = [Diagnostics.ProcessStartInfo]::new($Executable)
    $info.WorkingDirectory = $WorkingDirectory
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $info.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::Start($info)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $failure = $null
    try {
        @{ processId = $process.Id; startedAtUtcFileTime = $process.StartTime.ToUniversalTime().ToFileTimeUtc() } |
            ConvertTo-Json | Set-Content -LiteralPath "$LogBase.process.json"
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $timedOut = -not $process.WaitForExit($DeadlineSeconds * 1000)
        if ($timedOut) {
            $process.Kill($true)
            $process.WaitForExit()
        }
        $output = $stdout.GetAwaiter().GetResult()
        $errorOutput = $stderr.GetAwaiter().GetResult()
        $clock.Stop()
        $output | Set-Content -LiteralPath "$LogBase.stdout.txt"
        $errorOutput | Set-Content -LiteralPath "$LogBase.stderr.txt"
        if ($timedOut) { throw "Hard deadline exceeded. See $LogBase.*." }
        return [pscustomobject]@{
            exitCode = $process.ExitCode
            output = $output
            elapsedSeconds = [math]::Round($clock.Elapsed.TotalSeconds, 3)
        }
    }
    catch {
        $failure = $_.Exception
        throw
    }
    finally {
        try {
            if (-not $process.HasExited) {
                $process.Kill($true)
                $process.WaitForExit()
            }
        }
        catch {
            if ($null -ne $failure) {
                throw [AggregateException]::new('Command and owned-process cleanup both failed.',
                    [Exception[]]@($failure, $_.Exception))
            }
            throw
        }
        finally { $process.Dispose() }
    }
}

function Get-ExcelBehaviorDiscoveredCases {
    param([string]$Output)
    $lines = $Output -split '\r?\n'
    $header = [Array]::IndexOf($lines, 'The following Tests are available:')
    if ($header -lt 0) { throw 'Test discovery did not produce its case listing.' }
    foreach ($line in $lines[($header + 1)..($lines.Length - 1)]) {
        if ($line -match '^    (\S.*)$') { $Matches[1] }
    }
}

function Get-ExcelBehaviorReportCases {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing TRX report: $Path" }
    [xml]$trx = Get-Content -LiteralPath $Path -Raw
    return @($trx.TestRun.Results.UnitTestResult | ForEach-Object testName)
}

function Assert-ExcelBehaviorCases {
    param([string[]]$Expected, [string[]]$Actual)
    if ($Expected.Count -eq 0) { throw 'The required selection discovered zero cases.' }
    $counts = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal)
    foreach ($name in $Expected) {
        if (-not $counts.ContainsKey($name)) { $counts[$name] = 0 }
        $counts[$name]++
    }
    foreach ($name in $Actual) {
        if (-not $counts.ContainsKey($name) -or $counts[$name] -eq 0) {
            throw "Unexpected or duplicate case: $name"
        }
        $counts[$name]--
    }
    $missing = @($counts.GetEnumerator() | Where-Object Value -ne 0 | ForEach-Object Key)
    if ($missing.Count -ne 0) { throw "Discovered cases were omitted: $($missing -join ', ')" }
}

function Assert-ExcelBehaviorFullCoverage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Projects,
        [Parameter(Mandatory)][Collections.IDictionary]$ExpectedByProject,
        [Parameter(Mandatory)][Collections.IDictionary[]]$Stages,
        [Parameter(Mandatory)][string]$ResultsDirectory
    )

    foreach ($project in $Projects) {
        $actual = @(
            foreach ($stage in $Stages | Where-Object {
                $_.project -eq $project -and $_.name -ne 'ComInterop-infrastructure'
            }) {
                if ($stage.status -ne 'not-applicable-no-discovered-cases') {
                    Get-ExcelBehaviorReportCases -Path (Join-Path $ResultsDirectory "$($stage.name).trx")
                }
            }
        )
        Assert-ExcelBehaviorCases -Expected $ExpectedByProject[$project] -Actual $actual
    }
}

function Test-ExcelBehaviorStageRequiresDiscovery {
    [CmdletBinding()]
    param(
        [switch]$Full,
        [switch]$DiscoverOnly,
        [switch]$AllowEmpty
    )

    return (-not $Full) -or $DiscoverOnly -or $AllowEmpty
}

function Assert-ExcelBehaviorReport {
    param(
        [string]$Path,
        [AllowNull()][string[]]$Discovered,
        [Collections.IDictionary]$Evidence
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw "Missing TRX report: $Path" }
    [xml]$trx = Get-Content -LiteralPath $Path -Raw
    $summary = $trx.TestRun.ResultSummary
    $results = @($trx.TestRun.Results.UnitTestResult)
    $counters = $summary.Counters
    if ($null -eq $counters -or $null -eq $results[0]) { throw "Invalid TRX report: $Path" }
    $total = [int]::Parse($counters.total, [Globalization.CultureInfo]::InvariantCulture)
    $passed = [int]::Parse($counters.passed, [Globalization.CultureInfo]::InvariantCulture)
    $executed = [int]::Parse($counters.executed, [Globalization.CultureInfo]::InvariantCulture)
    if ($null -ne $Evidence) {
        $Evidence.total = $total
        $Evidence.passed = $passed
        $Evidence.failed = [int]::Parse($counters.failed, [Globalization.CultureInfo]::InvariantCulture)
        $Evidence.skipped = [int]::Parse($counters.notExecuted, [Globalization.CultureInfo]::InvariantCulture)
    }
    foreach ($name in @('failed', 'notExecuted', 'error', 'timeout', 'aborted',
        'inconclusive', 'passedButRunAborted', 'notRunnable', 'disconnected', 'warning',
        'inProgress', 'pending')) {
        if ($counters.HasAttribute($name)) {
            $count = [int]::Parse($counters.GetAttribute($name), [Globalization.CultureInfo]::InvariantCulture)
            if ($count -ne 0) { throw "Non-passing or contradictory TRX counter ${name}: $count. Report: $Path" }
        }
        elseif ($name -in @('failed', 'notExecuted')) {
            throw "Missing required TRX counter ${name}. Report: $Path"
        }
    }
    if ($null -ne $Discovered) {
        Assert-ExcelBehaviorCases -Expected $Discovered -Actual @($results | ForEach-Object testName)
    }
    if ($total -ne $results.Count -or $total -le 0 -or $passed -ne $total -or $executed -ne $total -or
        $summary.outcome -ne 'Completed' -or @($results | Where-Object outcome -ne 'Passed').Count -ne 0) {
        throw "Failed, skipped, incomplete, or contradictory TRX results: $Path"
    }
    return $total
}

function Assert-ExcelBehaviorStageResults {
    param([Collections.IDictionary[]]$Stages)
    if ($Stages.Count -eq 0) { throw 'No validation stages were recorded.' }
    $unfinished = @($Stages | Where-Object {
        $_.status -notin @('passed', 'not-applicable-no-discovered-cases')
    })
    if ($unfinished.Count -ne 0) {
        throw "Validation failed or did not finish: $(($unfinished | ForEach-Object { "$($_.name) ($($_.status))" }) -join ', ')."
    }
}

function Get-ExcelBehaviorBuildTarget {
    [CmdletBinding()]
    param(
        [switch]$Full,
        [string]$Project
    )

    if ($Full) { return 'Sbroenne.ExcelMcp.sln' }
    if ([string]::IsNullOrWhiteSpace($Project)) { throw 'Focused mode requires a test project.' }
    return "tests\ExcelMcp.$Project.Tests\ExcelMcp.$Project.Tests.csproj"
}

function Get-ExcelBehaviorRequiredFilter {
    param([Parameter(Mandatory)][string]$Filter)
    "RequiresExcel=true&($Filter)"
}

$runDirectory = $null
$summary = [ordered]@{
    status = 'incomplete'
    mode = if ($Full) { 'full' } else { 'focused' }
    commands = @()
    stages = @()
}
$commandTimings = [Collections.Generic.List[object]]::new()
$previousOwnership = $env:EXCELMCP_TEST_OWNERSHIP_DIRECTORY
$previousLanguage = $env:DOTNET_CLI_UI_LANGUAGE
try {
    if ([string]::IsNullOrWhiteSpace($ResultsDirectory)) {
        $ResultsDirectory = Join-Path $root 'TestResults\ExcelIntegration'
    }
    $runDirectory = Join-Path ([IO.Path]::GetFullPath($ResultsDirectory)) ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $runDirectory | Out-Null
    Write-Host "Evidence: $runDirectory"
    $env:EXCELMCP_TEST_OWNERSHIP_DIRECTORY = Join-Path $runDirectory 'ownership'
    $env:DOTNET_CLI_UI_LANGUAGE = 'en'
    $buildTarget = Get-ExcelBehaviorBuildTarget -Full:$Full -Project $Project
    Write-Host "Building $buildTarget"
    $build = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds 1800 `
        -LogBase (Join-Path $runDirectory 'build') -Arguments @(
            'build', $buildTarget, '-c', 'Release', '--no-restore',
            '--disable-build-servers', '--verbosity', 'minimal', '-warnaserror')
    if ($build.exitCode -ne 0) { throw "Release build failed. See $runDirectory\build.*." }
    $commandTimings.Add([pscustomobject]@{
        name = 'build'; elapsedSeconds = $build.elapsedSeconds
    })
    Write-Host "Build completed in $($build.elapsedSeconds) seconds"

    $projects = @('Core', 'Service', 'CLI', 'McpServer', 'ComInterop')
    $stages = @()
    if ($Full) {
        $base = Get-ExcelBehaviorRequiredFilter 'RunType!=OnDemand'
        foreach ($name in $projects) {
            $mainFilter = switch ($name) {
                Core { "$base&Feature!=VBA&Feature!=VBATrust&Feature!=Screenshot" }
                Service { "$base&Feature!=VBA&Feature!=Screenshot" }
                { $_ -in 'CLI', 'McpServer' } { "$base&FullyQualifiedName!~VbaRun_OnMacroWorkbook" }
                default { $base }
            }
            $stages += [pscustomobject]@{ name = "$name-main"; project = $name; filter = $mainFilter; allowEmpty = $false }
        }
        $stages += [pscustomobject]@{
            name = 'ComInterop-infrastructure'; project = 'ComInterop'; allowEmpty = $false
            filter = Get-ExcelBehaviorRequiredFilter 'RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP'
        }
        foreach ($name in @('Core', 'Service', 'CLI', 'McpServer')) {
            $vba = switch ($name) {
                Core { '(Feature=VBA|Feature=VBATrust)' }
                Service { 'Feature=VBA' }
                default { 'FullyQualifiedName~VbaRun_OnMacroWorkbook' }
            }
            $stages += [pscustomobject]@{
                name = "$name-vba"; project = $name
                filter = "$base&$vba"; allowEmpty = $true
            }
        }
        foreach ($name in @('Core', 'Service')) {
            $stages += [pscustomobject]@{
                name = "$name-screenshot"; project = $name
                filter = "$base&Feature=Screenshot"; allowEmpty = $true
            }
        }
    }
    else {
        $stages = @([pscustomobject]@{
            name = "$Project-focused"; project = $Project
            filter = Get-ExcelBehaviorRequiredFilter $Filter; allowEmpty = $false
        })
    }

    foreach ($stage in $stages) {
        $discoverStage = Test-ExcelBehaviorStageRequiresDiscovery -Full:$Full `
            -DiscoverOnly:$DiscoverOnly -AllowEmpty:$stage.allowEmpty
        $cases = @()
        if ($discoverStage) {
            $projectPath = "tests\ExcelMcp.$($stage.project).Tests\ExcelMcp.$($stage.project).Tests.csproj"
            Write-Host "Discovering $($stage.name)"
            $discovery = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds 300 `
                -LogBase (Join-Path $runDirectory "$($stage.name)-discovery") -Arguments @(
                    'test', $projectPath, '-c', 'Release', '--no-build', '--no-restore',
                    '--disable-build-servers', '--list-tests', '--filter', $stage.filter)
            if ($discovery.exitCode -ne 0) { throw "Discovery failed for $($stage.name)." }
            $cases = @(Get-ExcelBehaviorDiscoveredCases -Output $discovery.output)
            $commandTimings.Add([pscustomobject]@{
                name = "$($stage.name)-discovery"; elapsedSeconds = $discovery.elapsedSeconds
            })
            Write-Host "Discovered $($cases.Count) cases for $($stage.name) in $($discovery.elapsedSeconds) seconds"
        }
        $stage | Add-Member NoteProperty cases $cases
        $stage | Add-Member NoteProperty discovered $discoverStage
        $record = [ordered]@{
            name = $stage.name; project = $stage.project; filter = $stage.filter
            discovered = if ($discoverStage) { $cases.Count } else { $null }
            reconciliation = if ($discoverStage) { 'stage-discovery' } else { 'project-execution-union' }
            status = 'not-run'
        }
        $summary.stages += $record
        if ($discoverStage -and $cases.Count -eq 0) {
            if (-not $stage.allowEmpty) { throw "Empty required selection: $($stage.name)." }
            $record.status = 'not-applicable-no-discovered-cases'
        }
    }

    $normalDiscoveredCases = @{}
    if ($Full) {
        foreach ($name in $projects) {
            Write-Host "Checking complete test inventory for $name"
            $all = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds 300 `
                -LogBase (Join-Path $runDirectory "$name-normal-discovery") -Arguments @(
                    'test', "tests\ExcelMcp.$name.Tests\ExcelMcp.$name.Tests.csproj",
                    '-c', 'Release', '--no-build', '--no-restore', '--list-tests',
                    '--filter', (Get-ExcelBehaviorRequiredFilter 'RunType!=OnDemand'))
            if ($all.exitCode -ne 0) { throw "Complete discovery failed for $name." }
            $expected = @(Get-ExcelBehaviorDiscoveredCases -Output $all.output)
            $commandTimings.Add([pscustomobject]@{
                name = "$name-normal-discovery"; elapsedSeconds = $all.elapsedSeconds
            })
            Write-Host "Found $($expected.Count) normal cases for $name in $($all.elapsedSeconds) seconds"
            $normalDiscoveredCases[$name] = $expected
            if ($DiscoverOnly) {
                $actual = @($stages | Where-Object { $_.project -eq $name -and $_.name -ne 'ComInterop-infrastructure' } |
                    ForEach-Object { $_.cases })
                Assert-ExcelBehaviorCases -Expected $expected -Actual $actual
            }
        }
    }

    if (-not $DiscoverOnly) {
        for ($index = 0; $index -lt $stages.Count; $index++) {
            $stage = $stages[$index]
            $record = $summary.stages[$index]
            if ($stage.allowEmpty -and $stage.cases.Count -eq 0) { continue }
            $record.status = 'running'
            $selectionEvidence = if ($stage.discovered) {
                "$($stage.cases.Count) discovered cases"
            } else {
                'cases reconciled against complete project execution'
            }
            Write-Host "Running $($stage.name): $selectionEvidence"
            $deadline = if ($Full -and $stage.project -eq 'Core' -and
                -not $PSBoundParameters.ContainsKey('StageTimeoutSeconds')) { 28800 } else { $StageTimeoutSeconds }
            $execution = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds $deadline `
                -LogBase (Join-Path $runDirectory $stage.name) -Arguments @(
                    'test', "tests\ExcelMcp.$($stage.project).Tests\ExcelMcp.$($stage.project).Tests.csproj",
                    '-c', 'Release', '--no-build', '--no-restore', '--disable-build-servers',
                    '--filter', $stage.filter, '--blame-hang-timeout', "${HangTimeoutSeconds}s",
                    '--results-directory', $runDirectory, '--logger', "trx;LogFileName=$($stage.name).trx")
            $record.exitCode = $execution.exitCode
            $record.wallSeconds = $execution.elapsedSeconds
            $commandTimings.Add([pscustomobject]@{
                name = $stage.name; elapsedSeconds = $execution.elapsedSeconds
            })
            try {
                $discoveredCases = if ($Full -and -not $DiscoverOnly -and -not $stage.discovered) {
                    $null
                } else {
                    $stage.cases
                }
                $record.executed = Assert-ExcelBehaviorReport -Path (Join-Path $runDirectory "$($stage.name).trx") `
                    -Discovered $discoveredCases -Evidence $record
                if ($execution.exitCode -ne 0) { throw "$($stage.name) failed, including possible assembly cleanup errors." }
                $record.status = 'passed'
                Write-Host "$($stage.name) completed in $($execution.elapsedSeconds) seconds"
            }
            catch {
                $record.status = 'failed'
                $record.error = $_.Exception.Message
                if (-not $ContinueOnFailure) { throw }
                Write-Warning "$($stage.name) failed; retaining its evidence and continuing the completed test runs."
            }
        }
        if ($Full) {
            Assert-ExcelBehaviorFullCoverage -Projects $projects -ExpectedByProject $normalDiscoveredCases `
                -Stages $summary.stages -ResultsDirectory $runDirectory
        }
        Assert-ExcelBehaviorStageResults -Stages $summary.stages
    }
    $summary.status = if ($DiscoverOnly) { 'discovery-only' } else { 'passed' }
    $global:LASTEXITCODE = 0
}
catch {
    $summary.status = 'failed'
    $summary.error = $_.Exception.Message
    foreach ($record in $summary.stages) {
        if ($record.status -eq 'running') { $record.status = 'failed' }
    }
    $global:LASTEXITCODE = 1
    throw
}
finally {
    if ($null -ne $runDirectory) {
        $summary.commands = @($commandTimings)
        $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runDirectory 'summary.json')
    }
    $env:EXCELMCP_TEST_OWNERSHIP_DIRECTORY = $previousOwnership
    $env:DOTNET_CLI_UI_LANGUAGE = $previousLanguage
}

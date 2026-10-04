<#
.SYNOPSIS
Runs focused or ordered Excel behavior validation and retains reconciled evidence.
.EXAMPLE
.\scripts\Test-ExcelBehavior.ps1 -Project Service -Filter 'Feature=Tables&RunType!=OnDemand'
.EXAMPLE
.\scripts\Test-ExcelBehavior.ps1 -Full
.EXAMPLE
.\scripts\Test-ExcelBehavior.ps1 -Full -ContinueOnFailure
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
        $output | Set-Content -LiteralPath "$LogBase.stdout.txt"
        $errorOutput | Set-Content -LiteralPath "$LogBase.stderr.txt"
        if ($timedOut) { throw "Hard deadline exceeded. See $LogBase.*." }
        return [pscustomobject]@{ exitCode = $process.ExitCode; output = $output }
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

function Assert-ExcelBehaviorReport {
    param([string]$Path, [string[]]$Discovered, [Collections.IDictionary]$Evidence)
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
    Assert-ExcelBehaviorCases -Expected $Discovered -Actual @($results | ForEach-Object testName)
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

function Get-ExcelBehaviorSourceIdentity {
    $paths = & git -C $root ls-files --cached --others --exclude-standard
    if ($LASTEXITCODE -ne 0) { throw 'Could not enumerate source identity.' }
    $manifest = @(
        foreach ($path in ($paths | Sort-Object -Unique)) {
            $absolute = Join-Path $root $path
            if (Test-Path -LiteralPath $absolute -PathType Leaf) {
                "$path`t$((Get-FileHash -LiteralPath $absolute -Algorithm SHA256).Hash)"
            }
            else { "$path`tDELETED" }
        }
    )
    $bytes = [Text.Encoding]::UTF8.GetBytes($manifest -join "`n")
    return [pscustomobject]@{
        sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
        files = $manifest
    }
}

$runDirectory = $null
$summary = [ordered]@{ status = 'incomplete'; mode = if ($Full) { 'full' } else { 'focused' }; stages = @() }
$previousOwnership = $env:EXCELMCP_TEST_OWNERSHIP_DIRECTORY
$previousLanguage = $env:DOTNET_CLI_UI_LANGUAGE
try {
    if ([string]::IsNullOrWhiteSpace($ResultsDirectory)) {
        $ResultsDirectory = Join-Path $root 'TestResults\ExcelBehavior'
    }
    $runDirectory = Join-Path ([IO.Path]::GetFullPath($ResultsDirectory)) ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $runDirectory | Out-Null
    Write-Host "Evidence: $runDirectory"
    $env:EXCELMCP_TEST_OWNERSHIP_DIRECTORY = Join-Path $runDirectory 'ownership'
    $env:DOTNET_CLI_UI_LANGUAGE = 'en'
    $source = Get-ExcelBehaviorSourceIdentity
    $source | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runDirectory 'source.json')
    $summary.sourceSha256 = $source.sha256
    $summary.commit = (& git -C $root rev-parse HEAD)
    if ($LASTEXITCODE -ne 0) { throw 'Could not read the source commit.' }
    $build = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds 1800 `
        -LogBase (Join-Path $runDirectory 'build') -Arguments @(
            'build', 'Sbroenne.ExcelMcp.sln', '-c', 'Release', '--no-restore',
            '--disable-build-servers', '--verbosity', 'minimal', '-warnaserror')
    if ($build.exitCode -ne 0) { throw "Release build failed. See $runDirectory\build.*." }

    $projects = @('Core', 'Service', 'CLI', 'McpServer', 'ComInterop', 'SkillGeneration', 'ScriptSafety', 'Packaging', 'Diagnostics')
    $stages = @()
    if ($Full) {
        foreach ($name in $projects) {
            $base = 'RunType!=OnDemand'
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
            filter = 'RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP'
        }
        foreach ($name in @('Core', 'Service', 'CLI', 'McpServer')) {
            $vba = switch ($name) {
                Core { '(Feature=VBA|Feature=VBATrust)' }
                Service { 'Feature=VBA' }
                default { 'FullyQualifiedName~VbaRun_OnMacroWorkbook' }
            }
            $stages += [pscustomobject]@{
                name = "$name-vba"; project = $name
                filter = "RunType!=OnDemand&$vba"; allowEmpty = $true
            }
        }
        foreach ($name in @('Core', 'Service')) {
            $stages += [pscustomobject]@{
                name = "$name-screenshot"; project = $name
                filter = 'RunType!=OnDemand&Feature=Screenshot'; allowEmpty = $true
            }
        }
    }
    else {
        $stages = @([pscustomobject]@{ name = "$Project-focused"; project = $Project; filter = $Filter; allowEmpty = $false })
    }

    $binaryIdentity = @(
        foreach ($name in ($stages.project | Sort-Object -Unique)) {
            $assembly = @(Get-ChildItem -LiteralPath (Join-Path $root "tests\ExcelMcp.$name.Tests\bin\Release") `
                -Recurse -Filter "Sbroenne.ExcelMcp.$name.Tests.dll")
            if ($assembly.Count -ne 1) { throw "Ambiguous or missing Release test assembly for $name." }
            @{ project = $name; sha256 = (Get-FileHash -LiteralPath $assembly[0].FullName).Hash }
        }
    )
    $binaryIdentity | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runDirectory 'binaries.json')
    foreach ($stage in $stages) {
        $projectPath = "tests\ExcelMcp.$($stage.project).Tests\ExcelMcp.$($stage.project).Tests.csproj"
        $discovery = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds 300 `
            -LogBase (Join-Path $runDirectory "$($stage.name)-discovery") -Arguments @(
                'test', $projectPath, '-c', 'Release', '--no-build', '--no-restore',
                '--disable-build-servers', '--list-tests', '--filter', $stage.filter)
        if ($discovery.exitCode -ne 0) { throw "Discovery failed for $($stage.name)." }
        $cases = @(Get-ExcelBehaviorDiscoveredCases -Output $discovery.output)
        $stage | Add-Member NoteProperty cases $cases
        $record = [ordered]@{
            name = $stage.name; project = $stage.project; filter = $stage.filter
            discovered = $cases.Count; status = 'not-run'
        }
        $summary.stages += $record
        if ($cases.Count -eq 0) {
            if (-not $stage.allowEmpty) { throw "Empty required selection: $($stage.name)." }
            $record.status = 'not-applicable-no-discovered-cases'
        }
    }

    if ($Full) {
        foreach ($name in $projects) {
            $all = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds 300 `
                -LogBase (Join-Path $runDirectory "$name-normal-discovery") -Arguments @(
                    'test', "tests\ExcelMcp.$name.Tests\ExcelMcp.$name.Tests.csproj",
                    '-c', 'Release', '--no-build', '--no-restore', '--list-tests', '--filter', 'RunType!=OnDemand')
            if ($all.exitCode -ne 0) { throw "Complete discovery failed for $name." }
            $expected = @(Get-ExcelBehaviorDiscoveredCases -Output $all.output)
            $actual = @($stages | Where-Object { $_.project -eq $name -and $_.name -ne 'ComInterop-infrastructure' } |
                ForEach-Object { $_.cases })
            Assert-ExcelBehaviorCases -Expected $expected -Actual $actual
        }
    }

    if ((Get-ExcelBehaviorSourceIdentity).sha256 -ne $source.sha256) {
        throw 'Source changed during build/discovery; evidence does not match the current source.'
    }
    if (-not $DiscoverOnly) {
        for ($index = 0; $index -lt $stages.Count; $index++) {
            $stage = $stages[$index]
            $record = $summary.stages[$index]
            if ($stage.cases.Count -eq 0) { continue }
            $record.status = 'running'
            Write-Host "Running $($stage.name): $($stage.cases.Count) discovered cases"
            $deadline = if ($Full -and $stage.project -eq 'Core' -and
                -not $PSBoundParameters.ContainsKey('StageTimeoutSeconds')) { 28800 } else { $StageTimeoutSeconds }
            $execution = Invoke-ExcelBehaviorProcess -WorkingDirectory $root -DeadlineSeconds $deadline `
                -LogBase (Join-Path $runDirectory $stage.name) -Arguments @(
                    'test', "tests\ExcelMcp.$($stage.project).Tests\ExcelMcp.$($stage.project).Tests.csproj",
                    '-c', 'Release', '--no-build', '--no-restore', '--disable-build-servers',
                    '--filter', $stage.filter, '--blame-hang-timeout', "${HangTimeoutSeconds}s",
                    '--results-directory', $runDirectory, '--logger', "trx;LogFileName=$($stage.name).trx")
            $record.exitCode = $execution.exitCode
            try {
                $record.executed = Assert-ExcelBehaviorReport -Path (Join-Path $runDirectory "$($stage.name).trx") `
                    -Discovered $stage.cases -Evidence $record
                if ($execution.exitCode -ne 0) { throw "$($stage.name) failed, including possible assembly cleanup errors." }
                $record.status = 'passed'
            }
            catch {
                $record.status = 'failed'
                $record.error = $_.Exception.Message
                if (-not $ContinueOnFailure) { throw }
                Write-Warning "$($stage.name) failed; retaining its evidence and continuing the completed test runs."
            }
        }
        if ((Get-ExcelBehaviorSourceIdentity).sha256 -ne $source.sha256) {
            throw 'Source changed during execution; rerun against the final source.'
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
        $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runDirectory 'summary.json')
    }
    $env:EXCELMCP_TEST_OWNERSHIP_DIRECTORY = $previousOwnership
    $env:DOTNET_CLI_UI_LANGUAGE = $previousLanguage
}

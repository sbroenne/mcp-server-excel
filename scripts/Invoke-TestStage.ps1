#Requires -Version 7.4
. (Join-Path $PSScriptRoot 'Invoke-NativeCommand.ps1')

function Assert-TestReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string[]]$ExpectedCases)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing test report: $Path." }
    $document = [xml](Get-Content -LiteralPath $Path -Raw)
    $summary = $document.SelectSingleNode('//*[local-name()="ResultSummary"]')
    $counters = $summary.SelectSingleNode('*[local-name()="Counters"]')
    $results = @($document.SelectNodes('//*[local-name()="UnitTestResult"]'))
    $total = 0
    if (-not $counters -or -not [int]::TryParse($counters.GetAttribute('total'), [ref]$total) -or
        $total -le 0 -or $counters.GetAttribute('passed') -cne "$total" -or
        $counters.GetAttribute('executed') -cne "$total" -or $summary.GetAttribute('outcome') -cne 'Completed' -or
        $results.Count -ne $total -or @($results | Where-Object { $_.GetAttribute('outcome') -cne 'Passed' }).Count -or
        $summary.SelectNodes('.//*[local-name()="RunInfo"][@outcome="Error"]').Count) {
        throw "Test selection was empty, skipped, failed, or had a cleanup failure. See $Path."
    }
    foreach ($name in @('failed', 'notExecuted', 'error', 'timeout', 'aborted', 'inconclusive',
            'passedButRunAborted', 'notRunnable', 'disconnected', 'warning', 'inProgress', 'pending')) {
        $value = $counters.GetAttribute($name)
        if (($name -in @('failed', 'notExecuted') -or $value) -and $value -cne '0') {
            throw "Test selection has invalid $name count. See $Path."
        }
    }
    if ($PSBoundParameters.ContainsKey('ExpectedCases')) {
        if (-not $ExpectedCases.Count) { throw "Test case reconciliation failed: discovery was empty. See $Path." }
        $expected = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal)
        foreach ($name in $ExpectedCases) {
            if (-not $expected.ContainsKey($name)) { $expected[$name] = 0 }
            $expected[$name]++
        }
        foreach ($result in $results) {
            $name = $result.GetAttribute('testName')
            if (-not $expected.ContainsKey($name) -or $expected[$name] -eq 0) {
                throw "Test case reconciliation failed: unexpected or duplicate '$name'. See $Path."
            }
            $expected[$name]--
        }
        $missing = @($expected.Keys | Where-Object { $expected[$_] -gt 0 })
        if ($missing.Count) { throw "Test case reconciliation failed: omitted $($missing -join ', '). See $Path." }
    }
    return $total
}

function Get-DiscoveredTestCases([string]$Output) {
    $lines = $Output -split '\r?\n'
    $header = [Array]::IndexOf($lines, 'The following Tests are available:')
    if ($header -lt 0) { throw 'Test discovery output has no test-list header.' }
    $cases = @(
        for ($i = $header + 1; $i -lt $lines.Count -and $lines[$i].StartsWith('    ') -and $lines[$i].Trim(); $i++) {
            $lines[$i].Substring(4).TrimEnd()
        }
    )
    if (-not $cases.Count) { throw 'Test discovery returned no cases.' }
    return $cases
}

function Invoke-TestStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Project,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Filter,
        [Parameter(Mandatory)][string]$ResultsDirectory,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9_-]+$')][string]$Name,
        [ValidateRange(1, 28800)][int]$DeadlineSeconds = 1800,
        [string]$HangTimeout = '5m',
        [hashtable]$Environment = @{},
        [switch]$ListTests,
        [switch]$ReconcileCases
    )
    $root = Split-Path -Parent $PSScriptRoot
    $projectPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Project)
    $resultsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)
    if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf) -or [IO.Path]::GetExtension($projectPath) -notin @('.csproj', '.proj')) {
        throw "Missing or invalid test project: $projectPath."
    }
    $report = Join-Path $resultsPath "$Name.trx"
    if (Test-Path -LiteralPath $report) { throw "Use a new report path: $report." }
    New-Item -ItemType Directory -Path $resultsPath -Force | Out-Null
    $expected = $null
    if ($ReconcileCases -and -not $ListTests) {
        Invoke-TestStage -Project $projectPath -Filter $Filter -ResultsDirectory $resultsPath -Name "$Name-discovery" `
            -DeadlineSeconds $DeadlineSeconds -HangTimeout $HangTimeout -Environment $Environment -ListTests
        $expected = @(Get-DiscoveredTestCases (Get-Content (Join-Path $resultsPath "$Name-discovery.stdout.log") -Raw))
    }
    $settings = Join-Path $resultsPath "$Name.runsettings"
    $runSettings = [xml](Get-Content -LiteralPath (Join-Path $root 'tests\strict.runsettings') -Raw)
    $timeout = $runSettings.CreateElement('TestSessionTimeout')
    $timeout.InnerText = [string]($DeadlineSeconds * 1000)
    $null = $runSettings.RunSettings.RunConfiguration.AppendChild($timeout)
    $runSettings.Save($settings)
    $arguments = @('test', $projectPath, '-c', 'Release', '--no-build', '--no-restore', '--disable-build-servers',
        '-maxcpucount:1',
        '--settings', $settings, '--filter', $Filter, '--blame-hang-timeout', $HangTimeout,
        '--results-directory', $resultsPath, '--logger', "trx;LogFileName=$Name.trx")
    if ($ListTests) { $arguments += '--list-tests' } else {
        $arguments += @('--diag', (Join-Path $resultsPath "$Name.vstest.log"), '--logger', 'console;verbosity=normal')
    }
    $stageEnvironment = @{
        DOTNET_CLI_UI_LANGUAGE = 'en'
        DOTNET_CLI_USE_MSBUILD_SERVER = '0'
        MSBUILDDISABLENODEREUSE = '1'
    }
    foreach ($key in $Environment.Keys) { $stageEnvironment[$key] = $Environment[$key] }
    Write-Host "$Name : $Filter"
    $outputPath = Join-Path $resultsPath "$Name.stdout.log"
    $errorPath = Join-Path $resultsPath "$Name.stderr.log"
    Write-Host "Live logs: $outputPath ; $errorPath"
    try {
        $result = Invoke-NativeCommand -Executable dotnet -Arguments $arguments -DeadlineSeconds $DeadlineSeconds `
            -Environment $stageEnvironment -OutputPath $outputPath -ErrorPath $errorPath -ProgressName $Name -AllowFailure
    } catch [TimeoutException] {
        $message = "$Name exceeded the $DeadlineSeconds-second hard deadline. Evidence: $resultsPath."
        Set-Content -LiteralPath (Join-Path $resultsPath "$Name.deadline.log") -Value "$message`n$_"
        throw [TimeoutException]::new($message, $_.Exception)
    }
    Write-Host $result.Output
    if ($result.Error) { [Console]::Error.WriteLine($result.Error) }
    if ($result.ExitCode -ne 0) { throw "$Name failed with exit code $($result.ExitCode). Evidence: $resultsPath.`n$($result.Output)`n$($result.Error)" }
    if (-not $ListTests) {
        $check = @{ Path = $report }
        if ($ReconcileCases) { $check.ExpectedCases = $expected }
        $count = Assert-TestReport @check
        Write-Host "${Name}: $count cases passed."
    }
    $global:LASTEXITCODE = 0
}

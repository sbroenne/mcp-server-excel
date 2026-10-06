function Assert-TestReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing test report: $Path" }
    [xml]$trx = Get-Content -LiteralPath $Path -Raw
    $counters = $trx.TestRun.ResultSummary.Counters
    $results = @($trx.TestRun.Results.UnitTestResult)
    if (-not $counters -or [int]$counters.total -le 0 -or
        [int]$counters.passed -ne [int]$counters.total -or
        [int]$counters.executed -ne [int]$counters.total -or
        $results.Count -ne [int]$counters.total -or
        @($results | Where-Object outcome -ne 'Passed').Count -gt 0 -or
        $trx.TestRun.ResultSummary.outcome -ne 'Completed') {
        throw "Test selection was empty, skipped, failed, or had a cleanup failure. See $Path."
    }
    foreach ($name in @('failed', 'notExecuted', 'error', 'timeout', 'aborted',
        'inconclusive', 'passedButRunAborted', 'notRunnable', 'disconnected',
        'warning', 'inProgress', 'pending')) {
        if (($name -in @('failed', 'notExecuted') -and -not $counters.HasAttribute($name)) -or
            ($counters.HasAttribute($name) -and [int]$counters.GetAttribute($name) -ne 0)) {
            throw "Missing or non-passing test counter ${name}. See $Path."
        }
    }
}

function Invoke-TestStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Project,
        [Parameter(Mandatory)][string]$Filter,
        [Parameter(Mandatory)][string]$ResultsDirectory,
        [Parameter(Mandatory)][string]$Name,
        [ValidateRange(1, 28800)][int]$DeadlineSeconds = 1800,
        [string]$HangTimeout = '5m',
        [hashtable]$Environment = @{},
        [switch]$ListTests
    )

    $ResultsDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)
    $root = Split-Path -Parent $PSScriptRoot
    New-Item -ItemType Directory -Path $ResultsDirectory -Force | Out-Null
    $report = Join-Path $ResultsDirectory "$Name.trx"
    if (Test-Path -LiteralPath $report) { throw "Use a new report path: $report" }
    $info = [Diagnostics.ProcessStartInfo]::new('dotnet')
    $info.WorkingDirectory = $root
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.Environment['EXCELMCP_TEST_OWNERSHIP_DIRECTORY'] = Join-Path $ResultsDirectory "$Name-ownership"
    foreach ($key in $Environment.Keys) { $info.Environment[$key] = [string]$Environment[$key] }
    foreach ($argument in @('test', $Project, '-c', 'Release', '--no-build', '--no-restore',
        '--disable-build-servers', '--filter', $Filter, '--blame-hang-timeout', $HangTimeout,
        '--results-directory', $ResultsDirectory, '--logger', "trx;LogFileName=$Name.trx")) {
        $info.ArgumentList.Add($argument)
    }
    if ($ListTests) { $info.ArgumentList.Add('--list-tests') }
    Write-Host "$Name : $Filter"
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $process = [Diagnostics.Process]::Start($info)
    $stdout = $process.StandardOutput.ReadToEndAsync()
    $stderr = $process.StandardError.ReadToEndAsync()
    try {
        if (-not $process.WaitForExit($DeadlineSeconds * 1000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "$Name exceeded the $DeadlineSeconds-second hard deadline. Evidence: $ResultsDirectory"
        }
        if ($process.ExitCode -ne 0) { throw "$Name failed with exit code $($process.ExitCode). Evidence: $ResultsDirectory" }
    }
    finally {
        try {
            $output = $stdout.GetAwaiter().GetResult()
            $errors = $stderr.GetAwaiter().GetResult()
            $output | Set-Content -LiteralPath (Join-Path $ResultsDirectory "$Name.stdout.log") -Encoding utf8
            $errors | Set-Content -LiteralPath (Join-Path $ResultsDirectory "$Name.stderr.log") -Encoding utf8
            if ($output) { Write-Host $output }
            if ($errors) { Write-Host $errors }
        }
        finally {
            $process.Dispose()
            Write-Host "$Name wall time: $($clock.Elapsed)"
        }
    }
    if (-not $ListTests) { Assert-TestReport -Path $report }
    $global:LASTEXITCODE = 0
}

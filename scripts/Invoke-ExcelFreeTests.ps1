[CmdletBinding()]
param(
    [switch]$Local,
    [switch]$HookTests,
    [switch]$Contracts,
    [string[]]$ChangedPaths = @()
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$selections = [ordered]@{}
if ($Local) {
    if ($HookTests) { $selections['SkillGeneration'] = 'Feature=PreCommit|Feature=AutomationSafety' }
    if ($Contracts) {
        $selections['Core'] = 'Feature=GeneratedContracts'
        $selections['CLI'] = 'FullyQualifiedName~GeneratedActionContractCliTests'
        $selections['McpServer'] = 'FullyQualifiedName~CoreCommandsCoverageTests|FullyQualifiedName~McpToolSurfaceTests|FullyQualifiedName~CalculationGuidanceContractTests|FullyQualifiedName~GeneratedMcpParameterTests'
    }
    foreach ($path in $ChangedPaths) {
        if ($path -match '(PluginPublication|Publish-PreparedPlugins|PluginContent|AwesomeCopilotPolicy|Update-AwesomeCopilot|update-awesome-copilot|publish-plugins)') {
            $selections['SkillGeneration'] = if ($selections['SkillGeneration']) {
                "$($selections['SkillGeneration'])|Feature=PluginPublication"
            } else { 'Feature=PluginPublication' }
        }
        if ($path -match '^tests[/\\]ExcelMcp\.(Core|CLI|ComInterop|McpServer|Portable|Service)\.Tests[/\\]') {
            $selections[$Matches[1]] = 'RequiresExcel=false'
        }
    }
}
else {
    foreach ($project in @('CLI', 'ComInterop', 'Core', 'McpServer', 'Portable', 'Service', 'SkillGeneration')) {
        $selections[$project] = 'RequiresExcel=false'
    }
}
$results = Join-Path $root "TestResults\excel-free-$([Guid]::NewGuid().ToString('N'))"
foreach ($entry in $selections.GetEnumerator()) {
    if (-not [OperatingSystem]::IsWindows() -and $entry.Key -notin @('Portable', 'SkillGeneration')) {
        Write-Warning "$($entry.Key) tests target Microsoft.WindowsDesktop.App and cannot run on this host."
        continue
    }
    $project = Join-Path $root "tests\ExcelMcp.$($entry.Key).Tests\ExcelMcp.$($entry.Key).Tests.csproj"
    $filter = if ($entry.Key -eq 'Portable') {
        'RequiresExcel!=true&RunType!=OnDemand'
    } else {
        "RequiresExcel=false&RunType!=OnDemand&($($entry.Value))"
    }
    $info = [Diagnostics.ProcessStartInfo]::new('dotnet')
    $info.WorkingDirectory = $root
    $info.UseShellExecute = $false
    foreach ($argument in @('test', $project, '-c', 'Release', '--no-build', '--no-restore',
        '--filter', $filter, '--blame-hang-timeout', '5m',
        '--results-directory', $results, '--logger', "trx;LogFileName=$($entry.Key).trx")) {
        $info.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::Start($info)
    try {
        if (-not $process.WaitForExit(1800000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "$($entry.Key) tests exceeded the 30-minute deadline."
        }
        if ($process.ExitCode -ne 0) { throw "$($entry.Key) tests failed with exit code $($process.ExitCode)." }
    }
    finally { $process.Dispose() }
    $report = Join-Path $results "$($entry.Key).trx"
    if (-not (Test-Path -LiteralPath $report)) { throw "No test report for $($entry.Key)." }
    [xml]$trx = Get-Content -LiteralPath $report -Raw
    $counters = $trx.TestRun.ResultSummary.Counters
    if ([int]$counters.total -le 0 -or [int]$counters.passed -ne [int]$counters.total) {
        throw "$($entry.Key) selection was empty, skipped, or failed. See $report."
    }
}
$global:LASTEXITCODE = 0

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Acceptance', 'Editing', 'Reporting', 'Data', 'Lifecycle', 'Infrastructure', 'VBA', 'Desktop')]
    [string[]]$Groups = @('Editing', 'Reporting', 'Data', 'Lifecycle', 'Infrastructure', 'Acceptance', 'VBA', 'Desktop'),
    [switch]$IncludeInfrastructureDiagnostics,
    [switch]$ListTests,
    [string]$ResultsDirectory,
    [ValidateRange(1, 28800)][int]$DeadlineSeconds = 7200
)
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'Real-Excel validation requires Windows with desktop Excel.' }
if (-not $ListTests -and $null -eq [Type]::GetTypeFromProgID('Excel.Application')) {
    throw 'Desktop Excel is not registered. Selected Excel tests were not run.'
}
. (Join-Path $PSScriptRoot 'Invoke-TestStage.ps1')
. (Join-Path $PSScriptRoot 'Get-ExcelTestGroups.ps1')
$root = Split-Path -Parent $PSScriptRoot
if (-not $ResultsDirectory) { $ResultsDirectory = Join-Path $root "TestResults\excel-$([Guid]::NewGuid().ToString('N'))" }
$inventory = @(Get-ExcelTestGroups -ResultsDirectory (Join-Path $ResultsDirectory 'inventory'))
$pipe = "excelmcp-feature-$PID-$([Guid]::NewGuid().ToString('N'))"
$failures = [Collections.Generic.List[Exception]]::new()
try {
    foreach ($group in $Groups | Select-Object -Unique) {
        $selected = @($inventory | Where-Object {
            $_.Group -eq $group -and @($_.Cases | Where-Object { -not $_.OnDemand }).Count -gt 0
        })
        if (-not $selected) { throw "$group has no matching Excel classes." }
        if ($group -eq 'Acceptance' -and -not $ListTests) {
            & (Join-Path $PSScriptRoot 'Test-E2E.ps1') -SkipBuild -PipeName $pipe `
                -ResultsDirectory (Join-Path $ResultsDirectory 'acceptance')
            if ($LASTEXITCODE -ne 0) { throw 'Required acceptance stages failed.' }
            $selected = @(
                foreach ($item in $selected) {
                    $remaining = @($item.Cases | Where-Object { -not $_.Required -and -not $_.OnDemand })
                    if ($remaining.Count) {
                        $item.Filter = '(' + (($remaining | ForEach-Object { "FullyQualifiedName=$($_.Method)" }) -join '|') + ')'
                        $item
                    }
                }
            )
        }
        foreach ($project in $selected | Group-Object Project) {
            $filters = $project.Group.Filter -join '|'
            $filter = "RequiresExcel=true&RunType!=OnDemand&($filters)"
            Invoke-TestStage -Project $project.Name -Filter $filter -ResultsDirectory $ResultsDirectory `
                -Name "$group-$($project.Group[0].ProjectName)" -DeadlineSeconds $DeadlineSeconds `
                -HangTimeout 10m -ListTests:$ListTests -ReconcileCases -Environment @{ EXCELMCP_CLI_PIPE = $pipe }
        }
    }
    if ($IncludeInfrastructureDiagnostics) {
        $project = Join-Path $root 'tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj'
        Invoke-TestStage -Project $project -Filter 'RequiresExcel=true&RunType=OnDemand&FullyQualifiedName!~BeginBatch_RealIrmWorkbook&Locale!=ja-JP' `
            -ResultsDirectory $ResultsDirectory -Name Infrastructure-OnDemand -DeadlineSeconds 5400 `
            -HangTimeout 10m -ListTests:$ListTests -ReconcileCases -Environment @{ EXCELMCP_CLI_PIPE = $pipe }
    }
}
catch {
    $failures.Add($_.Exception)
}
finally {
    try {
        if (-not $ListTests) {
            & (Join-Path $PSScriptRoot 'Stop-ExcelCliService.ps1') -PipeName $pipe
            if ($LASTEXITCODE -ne 0) { throw 'Owned local-test CLI cleanup failed.' }
        }
    }
    catch { $failures.Add($_.Exception) }
}
if ($failures.Count) { throw [AggregateException]::new('Excel validation failed.', $failures) }
$global:LASTEXITCODE = 0

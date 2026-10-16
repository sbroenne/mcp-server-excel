#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Acceptance', 'Editing', 'Reporting', 'Data', 'Lifecycle', 'Infrastructure', 'VBA', 'Desktop')]
    [string[]]$Groups = @('Editing', 'Reporting', 'Data', 'Lifecycle', 'Infrastructure', 'Acceptance', 'VBA', 'Desktop'),
    [switch]$IncludeInfrastructureDiagnostics,
    [switch]$ListTests,
    [string]$PlanFile,
    [string]$ResultsDirectory,
    [ValidateRange(1, 28800)][int]$DeadlineSeconds = 7200
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Get-ExcelTestGroups.ps1')
if (-not $ResultsDirectory) { $ResultsDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) "TestResults\excel-$([Guid]::NewGuid().ToString('N'))" }
if ($PlanFile) {
    Invoke-ExcelPlan -Plan (Read-ValidationPlan $PlanFile) -ResultsDirectory $ResultsDirectory -ListTests:$ListTests -DeadlineSeconds $DeadlineSeconds
} else {
    if (-not $IsWindows) { throw 'Real-Excel validation requires Windows with desktop Excel.' }
    if (-not $ListTests -and -not [type]::GetTypeFromProgID('Excel.Application')) {
        throw 'Desktop Excel is not registered. Selected Excel tests were not run.'
    }
    $inventory = @(Get-ExcelTestGroups -ResultsDirectory (Join-Path $ResultsDirectory inventory))
    $plan = Get-ValidationPlan -Paths @()
    foreach ($group in @($Groups | Select-Object -Unique)) {
        $selected = @($inventory | Where-Object { $_.Group -eq $group -and @($_.Cases | Where-Object { -not $_.OnDemand }).Count })
        if (-not $selected.Count) { throw "$group has no matching Excel classes." }
        foreach ($item in $selected) {
            $filter = $item.Filter
            if ($group -eq 'Acceptance' -and -not $ListTests) {
                $plan.FullE2E = $true
                $methods = @($item.Cases | Where-Object { -not $_.Required -and -not $_.OnDemand })
                if (-not $methods.Count) { continue }
                $filter = '(' + (($methods.Method | ForEach-Object { "FullyQualifiedName=$_" }) -join '|') + ')'
            }
            $plan.ExcelSelections += @{ Project = $item.ProjectName; Filter = $filter; Area = $group }
        }
    }
    $plan.Excel = $true
    $plan.InfrastructureDiagnostics = [bool]$IncludeInfrastructureDiagnostics
    Invoke-ExcelPlan -Plan $plan -ResultsDirectory $ResultsDirectory -ListTests:$ListTests -DeadlineSeconds $DeadlineSeconds
}
$global:LASTEXITCODE = 0

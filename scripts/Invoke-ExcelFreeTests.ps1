#Requires -Version 7.4
[CmdletBinding()]
param(
    [switch]$Local, [switch]$HookTests, [switch]$Contracts, [switch]$SkillTests, [switch]$PackagingTests,
    [string[]]$ChangedPaths = @(),
    [ValidateSet('Commit', 'Fast', 'Process', 'Tooling')][string]$Group,
    [string]$PlanFile, [string]$ResultsDirectory
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ValidationExecution.ps1')
if ($PlanFile -and -not $Group -and -not $Local) { throw 'PlanFile requires an explicit Group or Local.' }
$plan = if ($PlanFile) { Read-ValidationPlan $PlanFile }
elseif ($Group) { Get-ValidationPlan -Full }
else { Get-ValidationPlan -Paths $ChangedPaths }
$selected = @(Get-FreeTestSelection -Plan $plan -Group $Group -Local:$Local -HookTests:$HookTests `
    -Contracts:$Contracts -SkillTests:$SkillTests -PackagingTests:$PackagingTests)
if (-not $ResultsDirectory) { $ResultsDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) "TestResults\excel-free-$([Guid]::NewGuid().ToString('N'))" }
foreach ($item in $selected) {
    Invoke-TestStage -Project (Get-TestProject $item.Owner) -Filter $item.Filter -ResultsDirectory $ResultsDirectory -Name $item.Owner
}
if (-not $selected.Count) { Write-Host 'No local Excel-free tests selected.' }
$global:LASTEXITCODE = 0

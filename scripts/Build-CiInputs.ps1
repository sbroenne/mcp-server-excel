#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PlanFile,
    [Parameter(Mandatory)][ValidateSet('Fast', 'Process', 'Tooling')][string]$Group,
    [switch]$ListProjects
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ValidationExecution.ps1')
$plan = Read-ValidationPlan $PlanFile
if ($ListProjects) { Get-SelectedBuildProjects $plan $Group }
else { Invoke-SelectedBuild $plan $Group }
$global:LASTEXITCODE = 0

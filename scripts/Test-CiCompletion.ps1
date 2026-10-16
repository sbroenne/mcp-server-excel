#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Detection,
    [Parameter(Mandatory)][string]$Tests,
    [Parameter(Mandatory)][string]$Packages,
    [Parameter(Mandatory)][string]$Npm,
    [Parameter(Mandatory)][string]$BuildOnly,
    [Parameter(Mandatory)][string]$SelectedTests,
    [Parameter(Mandatory)][string]$SelectedBuildOnly
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
Invoke-ExcelMcpBuild -Arguments @('complete') -OptionsParameter '--completion-options' -Options @{
    Detection = $Detection
    Checks = @(
        @{ Name = 'tests'; Selected = $SelectedTests; Result = $Tests },
        @{ Name = 'packages'; Selected = 'true'; Result = $Packages },
        @{ Name = 'npm'; Selected = 'true'; Result = $Npm },
        @{ Name = 'build-only'; Selected = $SelectedBuildOnly; Result = $BuildOnly }
    )
}

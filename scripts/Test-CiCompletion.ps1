[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Detection,
    [Parameter(Mandatory)][string]$Tests,
    [Parameter(Mandatory)][string]$Packages,
    [Parameter(Mandatory)][string]$Npm,
    [Parameter(Mandatory)][string]$SelectedTests,
    [Parameter(Mandatory)][string]$SelectedPackages,
    [Parameter(Mandatory)][string]$SelectedNpm
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
Invoke-ExcelMcpBuild -Arguments @('complete') -OptionsParameter '--completion-options' -Options @{
    Detection = $Detection
    Checks = @(
        @{ Name = 'tests'; Selected = $SelectedTests; Result = $Tests },
        @{ Name = 'packages'; Selected = $SelectedPackages; Result = $Packages },
        @{ Name = 'npm'; Selected = $SelectedNpm; Result = $Npm }
    )
}

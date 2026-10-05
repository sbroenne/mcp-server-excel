[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Detection,
    [Parameter(Mandatory)][string]$Tests,
    [Parameter(Mandatory)][string]$Packages,
    [Parameter(Mandatory)][string]$Npm,
    [Parameter(Mandatory)][string]$Lockfiles,
    [Parameter(Mandatory)][string]$SelectedTests,
    [Parameter(Mandatory)][string]$SelectedPackages,
    [Parameter(Mandatory)][string]$SelectedNpm,
    [Parameter(Mandatory)][string]$SelectedLockfiles
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
Invoke-ExcelMcpBuild -Arguments @('complete') -OptionsParameter '--completion-options' -Options @{
    Detection = $Detection
    Checks = @(
        @{ Name = 'tests'; Selected = $SelectedTests; Result = $Tests },
        @{ Name = 'packages'; Selected = $SelectedPackages; Result = $Packages },
        @{ Name = 'npm'; Selected = $SelectedNpm; Result = $Npm },
        @{ Name = 'lockfiles'; Selected = $SelectedLockfiles; Result = $Lockfiles }
    )
}

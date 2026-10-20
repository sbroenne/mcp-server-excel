#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Detection,
    [Parameter(Mandatory)][string]$Tests,
    [Parameter(Mandatory)][string]$Packages,
    [Parameter(Mandatory)][string]$Npm,
    [Parameter(Mandatory)][string]$BuildOnly,
    [Parameter(Mandatory)][string]$SelectedTests,
    [Parameter(Mandatory)][string]$SelectedBuildOnly,
    [Parameter(Mandatory)][string]$SelectedPackages,
    [Parameter(Mandatory)][string]$SelectedNpm
)
$ErrorActionPreference = 'Stop'
if ($Detection -cne 'success') { throw "Change detection did not succeed: $Detection." }
foreach ($check in @(
    @{ Name = 'tests'; Selected = $SelectedTests; Result = $Tests },
    @{ Name = 'packages'; Selected = $SelectedPackages; Result = $Packages },
    @{ Name = 'npm'; Selected = $SelectedNpm; Result = $Npm },
    @{ Name = 'build-only'; Selected = $SelectedBuildOnly; Result = $BuildOnly }
)) {
    if ($check.Selected -cnotin @('true', 'false')) { throw "Invalid selection flag for $($check.Name)." }
    $expected = if ($check.Selected -ceq 'true') { 'success' } else { 'skipped' }
    if ($check.Result -cne $expected) { throw "$($check.Name) returned '$($check.Result)', expected '$expected'." }
}
Write-Host 'All selected CI checks succeeded.'
$global:LASTEXITCODE = 0

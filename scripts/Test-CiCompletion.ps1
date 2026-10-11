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
if ($Detection -ne 'success') { throw "Change detection did not succeed: $Detection." }
foreach ($check in @(
    @{ Name = 'tests'; Selected = $SelectedTests; Result = $Tests },
    @{ Name = 'packages'; Selected = $SelectedPackages; Result = $Packages },
    @{ Name = 'npm'; Selected = $SelectedNpm; Result = $Npm }
)) {
    $expected = switch -CaseSensitive ($check.Selected) {
        'true' { 'success' }
        'false' { 'skipped' }
        default { throw "Invalid selection for $($check.Name): $($check.Selected)." }
    }
    if ($check.Result -ne $expected) {
        throw "$($check.Name): expected $expected, received $($check.Result)."
    }
}
Write-Host 'All selected CI checks succeeded.'
$global:LASTEXITCODE = 0

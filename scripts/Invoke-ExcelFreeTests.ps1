[CmdletBinding()]
param(
    [switch]$Local,
    [switch]$HookTests,
    [switch]$Contracts,
    [switch]$SkillTests,
    [switch]$PackagingTests,
    [string[]]$ChangedPaths = @(),
    [ValidateSet('Fast', 'Process', 'Tooling')][string]$Group,
    [string]$PlanFile,
    [string]$ResultsDirectory
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
$options = @{
    Local = [bool]$Local; HookTests = [bool]$HookTests; Contracts = [bool]$Contracts
    SkillTests = [bool]$SkillTests; PackagingTests = [bool]$PackagingTests
    ChangedPaths = @($ChangedPaths); Group = $Group
}
if ($PlanFile) { $options.PlanFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PlanFile) }
if ($ResultsDirectory) { $options.ResultsDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory) }
Invoke-ExcelMcpBuild -Arguments @('test-free') -OptionsParameter '--test-options' -Options $options

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
. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
if ($PlanFile) {
    $arguments = @('test', '--plan', $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PlanFile),
        '--group', 'Excel', '--no-build', '--deadline-seconds', "$DeadlineSeconds")
    if ($ListTests) { $arguments += '--list-tests' }
    if ($ResultsDirectory) { $arguments += @('--results-directory', $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)) }
    Invoke-ExcelMcpBuild -Arguments $arguments
    return
}
$options = @{
    Groups = @($Groups); IncludeInfrastructureDiagnostics = [bool]$IncludeInfrastructureDiagnostics
    ListTests = [bool]$ListTests; DeadlineSeconds = $DeadlineSeconds
}
if ($ResultsDirectory) { $options.ResultsDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory) }
Invoke-ExcelMcpBuild -Arguments @('test-excel') -OptionsParameter '--test-options' -Options $options

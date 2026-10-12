[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PlanFile,
    [Parameter(Mandatory)][ValidateSet('Fast', 'Process', 'Tooling')][string]$Group,
    [switch]$ListProjects
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
$arguments = @('build', '--plan', $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PlanFile), '--group', $Group)
if ($ListProjects) { $arguments += '--list-projects' }
Invoke-ExcelMcpBuild -Arguments $arguments

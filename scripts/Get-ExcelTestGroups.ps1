. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')
function Get-ExcelTestGroups {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ResultsDirectory)
    $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)
    $json = Invoke-ExcelMcpBuild -Arguments @('inventory', '--results-directory', $path)
    $json -join [Environment]::NewLine | ConvertFrom-Json
}

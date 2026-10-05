. (Join-Path $PSScriptRoot 'Invoke-BuildTool.ps1')

function Assert-TestReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    Invoke-ExcelMcpBuild -Arguments @('verify-report', '--report', $Path)
}

function Invoke-TestStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Project,
        [Parameter(Mandatory)][string]$Filter,
        [Parameter(Mandatory)][string]$ResultsDirectory,
        [Parameter(Mandatory)][string]$Name,
        [ValidateRange(1, 28800)][int]$DeadlineSeconds = 1800,
        [string]$HangTimeout = '5m',
        [hashtable]$Environment = @{},
        [switch]$ListTests
    )
    Invoke-ExcelMcpBuild -Arguments @('stage') -OptionsParameter '--stage-options' -Options @{
        Project = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Project)
        Filter = $Filter
        ResultsDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ResultsDirectory)
        Name = $Name; DeadlineSeconds = $DeadlineSeconds; HangTimeout = $HangTimeout
        Environment = $Environment; ListTests = [bool]$ListTests
    }
}

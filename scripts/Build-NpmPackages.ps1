[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')][string]$Component = 'McpServer',
    [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Version,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RuntimeExecutable,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
    Operation = 'Npm'; Component = $Component; Architecture = $Architecture
    Version = $Version; RuntimeExecutable = $RuntimeExecutable
    OutputDirectory = $OutputDirectory
}

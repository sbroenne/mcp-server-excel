[CmdletBinding()]
param(
    [ValidateSet('McpServer', 'Cli')][string]$Component = 'McpServer',
    [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64',
    [switch]$ArchiveOnly,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LauncherPackage,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RuntimePackage
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
    Operation = 'VerifyNpm'; Component = $Component; Architecture = $Architecture
    ArchiveOnly = [bool]$ArchiveOnly; LauncherPackage = $LauncherPackage
    RuntimePackage = $RuntimePackage
}

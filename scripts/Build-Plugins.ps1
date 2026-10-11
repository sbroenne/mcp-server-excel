<#
.SYNOPSIS
    Builds Agent Plugins from canonical templates and complete prepared skills.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = 'plugins',
    [string]$SkillsDirectory
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
    Operation = 'Plugins'; Version = $Version; OutputDirectory = $OutputDir
    SkillsDirectory = $SkillsDirectory
}

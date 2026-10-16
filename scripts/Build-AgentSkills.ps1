#Requires -Version 7.4
<#
.SYNOPSIS
    Generates CLI discovery and report-formatting skills.
#>
[CmdletBinding()]
param(
    [string]$OutputDir,
    [string]$Version,
    [switch]$GenerateOnly
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
    Operation = 'Skills'; Version = $Version; OutputDirectory = $OutputDir
    GenerateOnly = [bool]$GenerateOnly
}

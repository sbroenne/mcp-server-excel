#Requires -Version 7.4
<#
.SYNOPSIS
    Creates and verifies distributable packages. Never publishes them.
.DESCRIPTION
    Checks all components by default. Explicit runs may name components.
#>
[CmdletBinding()]
param(
    [ValidateSet('Cli', 'Mcp', 'Extension', 'Mcpb', 'Plugins')]
    [string[]]$Components = @('Cli', 'Mcp', 'Extension', 'Mcpb', 'Plugins'),
    [string]$Version,
    [string]$SkillsDirectory,
    [string]$McpRuntimeExecutable,
    [string]$CliRuntimeExecutable,
    [string]$OutputDirectory,
    [switch]$SkipExtensionTests
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
    Operation = 'Release'; Components = @($Components)
    Version = $Version; SkillsDirectory = $SkillsDirectory
    McpRuntimeExecutable = $McpRuntimeExecutable; CliRuntimeExecutable = $CliRuntimeExecutable
    OutputDirectory = $OutputDirectory; SkipExtensionTests = [bool]$SkipExtensionTests
}

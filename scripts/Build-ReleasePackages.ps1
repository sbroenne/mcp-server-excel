<#
.SYNOPSIS
    Creates and verifies selected distributable packages. Never publishes them.
.DESCRIPTION
    PRs use the shared changed-area plan. Explicit runs default to all components.
#>
[CmdletBinding()]
param(
    [ValidateSet('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins')]
    [string[]]$Components = @('Cli', 'Mcp', 'Extension', 'Mcpb', 'Skills', 'Plugins'),
    [string]$BaseRef,
    [string]$HeadRef = 'HEAD',
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
    Operation = 'Release'; Components = @($Components); BaseRef = $BaseRef
    HeadRef = $HeadRef; Version = $Version; SkillsDirectory = $SkillsDirectory
    McpRuntimeExecutable = $McpRuntimeExecutable; CliRuntimeExecutable = $CliRuntimeExecutable
    OutputDirectory = $OutputDirectory; SkipExtensionTests = [bool]$SkipExtensionTests
}

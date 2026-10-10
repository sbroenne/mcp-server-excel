<#
.SYNOPSIS
    Packages the direct-npx Windows MCP server configuration for Claude Desktop.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = './artifacts'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'scripts\PackageHelpers.ps1')
Invoke-TypedPackage -SourceRoot $root -Options @{
    Operation = 'Mcpb'; Version = $Version
    OutputDirectory = [IO.Path]::GetFullPath($OutputDir, $PSScriptRoot)
}

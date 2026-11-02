#Requires -Version 7.4
<#
.SYNOPSIS
    Builds launcher-only Agent Plugins from canonical templates.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = 'plugins'
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
$root = Split-Path $PSScriptRoot -Parent
$Version = Get-PackageVersion $Version
$output = [IO.Path]::GetFullPath($OutputDir, $root)
Assert-PackageOutputPath $output $root
if ((Test-PackagePathContains $root $output) -and $output -ne (Join-Path $root 'plugins') -and
    -not (Test-PackagePathContains (Join-Path $root 'artifacts') $output)) { throw "Unsafe plugin output directory: $output" }

function Assert-PreparedPlugin {
    param([string]$Name, [string]$Directory)
    $path = Join-Path $Directory 'plugin.json'
    $manifest = Read-PackageJson $path
    foreach ($field in $manifest.Keys) {
        if ($field -cnotin @('$schema', 'name', 'version', 'description', 'author', 'homepage', 'repository', 'license', 'keywords', 'extensions')) {
            throw "$path contains unsupported Agent Plugins 1.0 field '$field'."
        }
    }
    if ($manifest.'$schema' -cne 'https://agent-plugins.org/schemas/1.0.0/plugin.schema.json') { throw "$path must target https://agent-plugins.org/schemas/1.0.0/plugin.schema.json." }
    if ($manifest.name -cne $Name -or $manifest.version -cne $Version) { throw "$path has unexpected name or version; expected '$Name' '$Version'." }
    if ($manifest.repository -isnot [string]) { throw "$path repository must be a string." }
    if (Get-ChildItem -LiteralPath $Directory -Recurse -File -Filter 'install-global.ps1') { throw 'Global installation helpers are retired; use npx instead.' }
    if ([IO.File]::Exists((Join-Path $Directory '.mcp.json'))) { throw "Legacy MCP configuration is not permitted in Agent Plugins 1.0 packages: $Directory" }
    $skills = Join-Path $Directory 'skills'
    if ([IO.Directory]::Exists($skills)) { throw "Plugins must not bundle skills: $Directory" }
    $mcpPath = Join-Path $Directory 'mcp.json'
    if (-not [IO.File]::Exists($mcpPath)) { return }
    $mcp = Read-PackageJson $mcpPath
    if ($mcp.Count -ne 2 -or -not $mcp.Contains('$schema') -or -not $mcp.Contains('mcpServers')) { throw "$mcpPath must contain only '`$schema' and 'mcpServers'." }
    if ($mcp.'$schema' -cne 'https://agent-plugins.org/schemas/1.0.0/mcp.schema.json') { throw "$mcpPath has an unexpected MCP schema." }
    if ($mcp.mcpServers -isnot [Collections.IDictionary]) { throw "Missing object mcpServers: $mcpPath" }
    foreach ($server in $mcp.mcpServers.Keys) {
        $config = $mcp.mcpServers[$server]
        if ($config -isnot [Collections.IDictionary]) { throw "Malformed server '$server'." }
        foreach ($field in $config.Keys) {
            if ($field -cnotin @('type', 'command', 'args', 'env', 'cwd')) { throw "$mcpPath server '$server' contains unsupported fields." }
        }
        if ($config.type -cne 'stdio' -or $config.command -isnot [string] -or $config.command -match '\s') { throw "$mcpPath server '$server' must use type 'stdio' and a single executable command token." }
        if ($config.args -isnot [array]) { throw "Missing array args: $mcpPath" }
        foreach ($argument in $config.args) {
            if ($argument -isnot [string] -or $argument.Contains('{pluginDir}')) { throw "$mcpPath server '$server' still uses the legacy '{pluginDir}' placeholder or invalid arguments." }
        }
    }
}

Invoke-PackageStage 'plugins' {
    param($stage)
    foreach ($component in 'excel-cli', 'excel-mcp') {
        $destination = Join-Path $stage $component
        Copy-PackageItem (Join-Path $root ".github\plugins\$component") $destination
        $bin = Join-Path $destination 'bin'
        if ([IO.Directory]::Exists($bin)) {
            foreach ($file in [IO.Directory]::EnumerateFiles($bin, '*', [IO.SearchOption]::AllDirectories)) {
                if ([IO.Path]::GetExtension($file) -in @('.exe', '.dll', '.pdb') -or $file -match '\.(deps|runtimeconfig)\.json$') { Remove-PackageItem $file }
            }
        }
        $manifestPath = Join-Path $destination 'plugin.json'
        $manifest = Read-PackageJson $manifestPath
        $manifest.version = $Version
        Write-PackageJson $manifestPath $manifest
        [IO.File]::WriteAllText((Join-Path $destination 'version.txt'), $Version)
        Assert-PreparedPlugin $component $destination
    }
    [IO.Directory]::CreateDirectory($output) | Out-Null
    foreach ($name in 'excel-cli', 'excel-mcp') { Assert-PackageNoLinks (Join-Path $output $name) }
    foreach ($name in 'excel-cli', 'excel-mcp') { Install-PackageOutput (Join-Path $stage $name) (Join-Path $output $name) }
    $output
    '[ok] excel-mcp - npx config'
    '[ok] excel-cli - argument-safe npx wrapper'
}

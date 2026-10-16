#Requires -Version 7.4
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
if (-not $Version) { $Version = ([xml][IO.File]::ReadAllText((Join-Path $root 'Directory.Build.props'))).Project.PropertyGroup.Version | Select-Object -First 1 }
$Version = Get-PackageVersion $Version
$output = [IO.Path]::GetFullPath($OutputDir, $PSScriptRoot)
Assert-PackageOutputPath $output $root
if ($output -eq $PSScriptRoot) { throw "Unsafe package output directory: $output" }
Invoke-PackageStage 'mcpb' {
    param($stage)
    $manifest = Read-PackageJson (Join-Path $PSScriptRoot 'manifest.json')
    $manifest.version = $Version
    Write-PackageJson (Join-Path $stage 'manifest.json') $manifest
    foreach ($name in 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md') {
        $source = if ($name -in @('LICENSE', 'CHANGELOG.md')) { $root } else { $PSScriptRoot }
        Copy-PackageItem (Join-Path $source $name) (Join-Path $stage $name)
    }
    $payload = Join-Path $stage 'payload'
    [IO.Directory]::CreateDirectory($payload) | Out-Null
    foreach ($name in 'manifest.json', 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md') {
        Move-PackageItem (Join-Path $stage $name) (Join-Path $payload $name)
    }
    $archive = Join-Path $stage 'bundle.zip'
    [IO.Compression.ZipFile]::CreateFromDirectory($payload, $archive)
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        foreach ($name in 'manifest.json', 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md') {
            if (-not $zip.GetEntry($name)) { throw "MCPB is missing required metadata: $name" }
        }
    } finally { $zip.Dispose() }
    [IO.Directory]::CreateDirectory($output) | Out-Null
    $destination = Join-Path $output "excel-mcp-$Version.mcpb"
    Install-PackageOutput $archive $destination
    Install-PackageOutput (Join-Path $payload 'manifest.json') (Join-Path $output 'manifest.json')
    $destination
}

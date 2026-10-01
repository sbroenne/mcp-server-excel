<#
.SYNOPSIS
    Packages the direct-npx Windows MCP server configuration for Claude Desktop.
.DESCRIPTION
    Bundles metadata only. An available npx resolves the latest public ExcelMcp
    package at launch. Never runs from pre-commit.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = './artifacts'
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'McpbPackaging.ps1')
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'scripts\PackageHelpers.ps1')
if (-not $Version) {
    [xml]$props = Get-Content (Join-Path $root 'Directory.Build.props')
    $Version = $props.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
}
if ($Version -notmatch '^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$') { throw 'A valid package version is required.' }
$output = [IO.Path]::GetFullPath($OutputDir, $PSScriptRoot)
Assert-PackageOutputPath -Path $output -RepoRoot $root
if ($output -eq [IO.Path]::GetPathRoot($output) -or $output -eq $root -or $output -eq $PSScriptRoot) {
    throw "Unsafe package output directory: $output"
}
$ancestor = $output
while ($ancestor) {
    if ((Test-Path -LiteralPath $ancestor) -and
        ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw "Package output must not traverse a link: $ancestor"
    }
    $ancestor = Split-Path $ancestor -Parent
}
$stage = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpMcpb-$([Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $stage | Out-Null
try {
    $manifest = Get-Content (Join-Path $PSScriptRoot 'manifest.json') -Raw | ConvertFrom-Json
    $manifest.version = $Version
    $manifest | ConvertTo-Json -Depth 20 | Set-Content (Join-Path $stage 'manifest.json') -Encoding utf8
    foreach ($name in @('icon-512.png', 'README.md')) { Copy-Item (Join-Path $PSScriptRoot $name) $stage }
    foreach ($name in @('LICENSE', 'CHANGELOG.md')) { Copy-Item (Join-Path $root $name) $stage }
    $entries = @('manifest.json', 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md') |
        ForEach-Object { Join-Path $stage $_ }
    $archive = Join-Path $stage 'package.zip'
    Compress-Archive -LiteralPath $entries -DestinationPath $archive -CompressionLevel Optimal
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        foreach ($name in @('manifest.json', 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md')) {
            if (-not $zip.GetEntry($name)) { throw "MCPB is missing required metadata: $name" }
        }
    }
    finally { $zip.Dispose() }
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    $destination = Join-Path $output "excel-mcp-$Version.mcpb"
    Install-PackageOutput -Source $archive -Destination $destination
    Install-PackageOutput -Source (Join-Path $stage 'manifest.json') -Destination (Join-Path $output 'manifest.json')
    Write-Host "Created $destination"
}
finally { Remove-McpbStagingDirectory -Path $stage }

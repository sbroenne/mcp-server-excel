<#
.SYNOPSIS
    Packages a prepared Windows MCP server for Claude Desktop.
.DESCRIPTION
    Pass RuntimeExecutable to reuse a published server. Without it, this explicit
    manual command publishes the server first. Never runs from pre-commit.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDir = './artifacts',
    [string]$RuntimeExecutable
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
Assert-PackageOutputPath -Path $output -RepoRoot $root -Inputs @($RuntimeExecutable)
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
New-Item -ItemType Directory -Path (Join-Path $stage 'server') -Force | Out-Null
try {
    if (-not $RuntimeExecutable) {
        $publish = Join-Path $stage 'publish'
        Publish-PackageRuntime -Component Mcp -RepoRoot $root -Version $Version -OutputDirectory $publish
        $RuntimeExecutable = Join-Path $publish 'Sbroenne.ExcelMcp.McpServer.exe'
    }
    $exe = Join-Path $stage 'server\excel-mcp-server.exe'
    Copy-Item -LiteralPath $RuntimeExecutable -Destination $exe
    if ($IsWindows) {
        $reportedVersion = & $exe --version 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0 -or $reportedVersion -notmatch [regex]::Escape($Version)) {
            throw "Prepared MCP server does not report package version $Version. $reportedVersion"
        }
    }
    $manifest = Get-Content (Join-Path $PSScriptRoot 'manifest.json') -Raw | ConvertFrom-Json
    $manifest.version = $Version
    $manifest | ConvertTo-Json -Depth 20 | Set-Content (Join-Path $stage 'manifest.json') -Encoding utf8
    foreach ($name in @('icon-512.png', 'README.md')) { Copy-Item (Join-Path $PSScriptRoot $name) $stage }
    foreach ($name in @('LICENSE', 'CHANGELOG.md')) { Copy-Item (Join-Path $root $name) $stage }
    $entries = @('manifest.json', 'icon-512.png', 'README.md', 'LICENSE', 'CHANGELOG.md', 'server') |
        ForEach-Object { Join-Path $stage $_ }
    $archive = Join-Path $stage 'package.zip'
    Compress-Archive -LiteralPath $entries -DestinationPath $archive -CompressionLevel Optimal
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        if (-not $zip.GetEntry('server/excel-mcp-server.exe')) { throw 'MCPB is missing its server executable.' }
    }
    finally { $zip.Dispose() }
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    $destination = Join-Path $output "excel-mcp-$Version.mcpb"
    Install-PackageOutput -Source $archive -Destination $destination
    Install-PackageOutput -Source (Join-Path $stage 'manifest.json') -Destination (Join-Path $output 'manifest.json')
    Write-Host "Created $destination"
}
finally { Remove-McpbStagingDirectory -Path $stage }

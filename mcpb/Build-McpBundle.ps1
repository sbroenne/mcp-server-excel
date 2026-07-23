<#
.SYNOPSIS
    Creates the MCPB (MCP Bundle) package for Claude Desktop.

.DESCRIPTION
    This script builds the MCP Server as a self-contained Windows x64 executable
    and packages it as an .mcpb file for one-click installation in Claude Desktop.

.PARAMETER Version
    The version number for the package (e.g., "1.0.0"). If not specified,
    reads from Directory.Build.props.

.PARAMETER OutputDir
    The output directory for the MCPB package. Defaults to ./artifacts

.EXAMPLE
    .\Build-McpBundle.ps1
    Creates MCPB package with version from Directory.Build.props

.EXAMPLE
    .\Build-McpBundle.ps1 -Version "1.2.0"
    Creates MCPB package with specified version

.NOTES
    Requirements:
    - .NET 10 SDK
    - Windows x64

    Output:
    mcpb/artifacts/excel-mcp-{version}.mcpb

    Contents:
    ├── manifest.json
    ├── icon-512.png
    ├── README.md
    ├── LICENSE
    ├── CHANGELOG.md
    └── server/
        └── excel-mcp-server.exe
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$Version,

    [Parameter()]
    [string]$OutputDir = "./artifacts"
)

$ErrorActionPreference = "Stop"

# Get script and project directories
$McpbDir = $PSScriptRoot
$RootDir = Split-Path $McpbDir -Parent
$McpServerDir = Join-Path $RootDir "src/ExcelMcp.McpServer"

Write-Host "🏗️  Building MCPB (MCP Bundle) package..." -ForegroundColor Cyan
Write-Host ""

# Determine version
if (-not $Version) {
    $PropsFile = Join-Path $RootDir "Directory.Build.props"
    if (Test-Path $PropsFile) {
        $xml = [xml](Get-Content $PropsFile)
        $Version = $xml.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    }
    if (-not $Version) {
        $Version = "1.0.0"
    }
}
Write-Host "📋 Version: $Version" -ForegroundColor Green

# Create output directory (relative to mcpb directory)
$OutputDir = Join-Path $McpbDir $OutputDir
if (Test-Path $OutputDir) {
    Remove-Item -Recurse -Force $OutputDir
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Create temp staging directory
$StagingDir = Join-Path $OutputDir "staging"
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

Write-Host ""
Write-Host "📦 Publishing self-contained executable..." -ForegroundColor Yellow

# Build self-contained executable with inline publish settings
# Note: ReadyToRun=false keeps exe small (~15 MB vs 100+ MB)
# Note: NuGetAudit=false avoids network failures during vulnerability check
$PublishArgs = @(
    "publish"
    "$McpServerDir/ExcelMcp.McpServer.csproj"
    "-c", "Release"
    "-r", "win-x64"
    "--self-contained", "true"
    "-p:PublishSingleFile=true"
    "-p:IncludeNativeLibrariesForSelfExtract=true"
    "-p:PublishTrimmed=false"
    "-p:PublishReadyToRun=false"
    "-p:NuGetAudit=false"
    "-p:Version=$Version"
    "-o", $StagingDir
    "--verbosity", "quiet"
)

& dotnet @PublishArgs
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Publish failed!" -ForegroundColor Red
    exit 1
}

Write-Host "   ✓ Built Sbroenne.ExcelMcp.McpServer.exe" -ForegroundColor Green

# Create server subdirectory and rename exe to match manifest
$ServerDir = Join-Path $StagingDir "server"
New-Item -ItemType Directory -Path $ServerDir -Force | Out-Null
$FinalExePath = Join-Path $ServerDir "excel-mcp-server.exe"
Move-Item (Join-Path $StagingDir "Sbroenne.ExcelMcp.McpServer.exe") $FinalExePath -Force
Write-Host "   ✓ Renamed to server/excel-mcp-server.exe" -ForegroundColor Green

# Verify executable works
$VersionOutput = & $FinalExePath --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Executable verification failed!" -ForegroundColor Red
    exit 1
}
Write-Host "   ✓ Verified: $VersionOutput" -ForegroundColor Green

# Copy manifest.json and update version
$ManifestSrc = Join-Path $McpbDir "manifest.json"
$ManifestDst = Join-Path $StagingDir "manifest.json"
$ManifestContent = Get-Content $ManifestSrc -Raw
# Update all version fields in manifest
$ManifestContent = $ManifestContent -replace '"version":\s*"[\d\.]+"', "`"version`": `"$Version`""
Set-Content $ManifestDst $ManifestContent -NoNewline
Write-Host "   ✓ Copied manifest.json (version: $Version)" -ForegroundColor Green

# Copy icon from mcpb directory
$IconSrc = Join-Path $McpbDir "icon-512.png"
$IconDst = Join-Path $StagingDir "icon-512.png"
Copy-Item $IconSrc $IconDst -Force
Write-Host "   ✓ Copied icon-512.png" -ForegroundColor Green

# Copy README.md from mcpb directory (end-user documentation)
$ReadmeSrc = Join-Path $McpbDir "README.md"
$ReadmeDst = Join-Path $StagingDir "README.md"
Copy-Item $ReadmeSrc $ReadmeDst -Force
Write-Host "   ✓ Copied README.md" -ForegroundColor Green

# Copy LICENSE from root directory (required for MCPB submission)
$LicenseSrc = Join-Path $RootDir "LICENSE"
$LicenseDst = Join-Path $StagingDir "LICENSE"
Copy-Item $LicenseSrc $LicenseDst -Force
Write-Host "   ✓ Copied LICENSE" -ForegroundColor Green

# Copy CHANGELOG.md from root directory (recommended for MCPB submission)
$ChangelogSrc = Join-Path $RootDir "CHANGELOG.md"
$ChangelogDst = Join-Path $StagingDir "CHANGELOG.md"
Copy-Item $ChangelogSrc $ChangelogDst -Force
Write-Host "   ✓ Copied CHANGELOG.md" -ForegroundColor Green

# Note: Agent Skills are NOT included in MCPB bundle - Claude Desktop doesn't use them.
# Skills are only bundled with the VS Code extension (via chatSkills contribution point).

# Create mcpb file (zip with .mcpb extension)
$McpbFileName = "excel-mcp-$Version.mcpb"
$McpbPath = Join-Path $OutputDir $McpbFileName

Write-Host ""
Write-Host "📦 Creating MCPB bundle..." -ForegroundColor Yellow

# Get files/directories to include (manifest.json, icon, README, LICENSE, CHANGELOG at root, server/ directory with exe)
$FilesToZip = @(
    (Join-Path $StagingDir "manifest.json"),
    (Join-Path $StagingDir "icon-512.png"),
    (Join-Path $StagingDir "README.md"),
    (Join-Path $StagingDir "LICENSE"),
    (Join-Path $StagingDir "CHANGELOG.md"),
    (Join-Path $StagingDir "server")
)

# Remove .mcp directory if it exists (MCP registry metadata not needed in MCPB bundle)
$McpMetaDir = Join-Path $StagingDir ".mcp"
if (Test-Path $McpMetaDir) {
    Remove-Item -Recurse -Force $McpMetaDir
    Write-Host "   ✓ Removed .mcp directory (not needed in MCPB)" -ForegroundColor DarkGray
}

Compress-Archive -Path $FilesToZip -DestinationPath $McpbPath -Force
Write-Host "   ✓ Created $McpbFileName" -ForegroundColor Green

# Copy manifest to output dir for verification
Copy-Item $ManifestDst (Join-Path $OutputDir "manifest.json") -Force

# Clean up staging
Remove-Item -Recurse -Force $StagingDir

# Show results
$McpbSize = (Get-Item $McpbPath).Length / 1MB
Write-Host ""
Write-Host "✅ MCPB bundle created successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "📁 Output:" -ForegroundColor Cyan
Write-Host "   $McpbPath" -ForegroundColor White
Write-Host "   Size: $([math]::Round($McpbSize, 1)) MB" -ForegroundColor White
Write-Host ""
Write-Host "📋 Contents:" -ForegroundColor Cyan

# List mcpb contents
$McpbContents = [System.IO.Compression.ZipFile]::OpenRead($McpbPath)
try {
    foreach ($entry in $McpbContents.Entries) {
        $sizeKB = [math]::Round($entry.Length / 1KB, 1)
        Write-Host "   - $($entry.FullName) ($sizeKB KB)" -ForegroundColor White
    }
} finally {
    $McpbContents.Dispose()
}

Write-Host ""
Write-Host "🚀 Installation:" -ForegroundColor Cyan
Write-Host "   Double-click the .mcpb file to install in Claude Desktop" -ForegroundColor White
Write-Host "   Or drag-and-drop onto Claude Desktop window" -ForegroundColor White
Write-Host ""
Write-Host "📤 Distribution:" -ForegroundColor Cyan
Write-Host "   1. Upload $McpbFileName to GitHub release" -ForegroundColor White
Write-Host "   2. Users can download and double-click to install" -ForegroundColor White
Write-Host "   3. Submit to Anthropic Directory for discoverability" -ForegroundColor White
Write-Host ""

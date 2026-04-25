<#
.SYNOPSIS
    Builds Copilot CLI plugins by copying validated published plugin templates and updating versions.

.DESCRIPTION
    Phase 3 build script that preserves Phase 1/2 validated plugin implementations.

    COPY STRATEGY (not generate):
    1. Copy validated plugin structure from the published marketplace repo (../mcp-server-excel-plugins/plugins/)
    2. Update version in plugin.json and version.txt
    3. Apply source-owned plugin overlays from .github/plugins/ (overlay content only, not standalone plugin roots)
    4. Refresh skills content from source repo
    5. Refresh shared references from source repo

    WHY COPY NOT GENERATE:
    - Phase 1/2 created validated bin scripts, READMEs, configs
    - Regenerating would introduce drift and regressions
    - Build script's job: version injection + source-owned overlays + skill refresh, not plugin authoring

    OUTPUT:
    plugins/
      excel-mcp/     → Full MCP plugin (copied structure + updated version + fresh skills)
      excel-cli/     → Full CLI plugin (copied structure + updated version + fresh skills)

.PARAMETER Version
    Plugin version. If not specified, reads from skills/excel-mcp/VERSION file.

.PARAMETER OutputDir
    Output directory. Default: plugins/

.PARAMETER PluginTemplateDir
    Validated plugin templates source. Default: ../mcp-server-excel-plugins/plugins/ in the published marketplace repo.

.EXAMPLE
    ./Build-Plugins.ps1

.EXAMPLE
    ./Build-Plugins.ps1 -Version 1.2.3
#>
param(
    [string]$Version = $null,
    [string]$OutputDir = "plugins",
    [string]$PluginTemplateDir = $null
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$SharedDir = Join-Path $SkillsDir "shared"
$PluginOverlayDir = Join-Path $RepoRoot ".github\plugins"

# Default template dir: sibling repo
if (-not $PluginTemplateDir) {
    $PluginTemplateDir = Join-Path (Split-Path -Parent $RepoRoot) "mcp-server-excel-plugins\plugins"
}

# Validate template directory exists
if (-not (Test-Path $PluginTemplateDir)) {
    Write-Error @"
❌ Plugin template directory not found: $PluginTemplateDir

Expected: ../mcp-server-excel-plugins/plugins/
This directory contains the published marketplace repo's validated plugin implementations.

If running in CI/CD, clone the published repo first:
  git clone https://github.com/sbroenne/mcp-server-excel-plugins ../mcp-server-excel-plugins
"@
    exit 1
}

# Resolve version
if (-not $Version) {
    $VersionFile = Join-Path $SkillsDir "excel-mcp\VERSION"
    if (Test-Path $VersionFile) {
        $Version = (Get-Content $VersionFile -Raw).Trim()
        Write-Host "Using version from VERSION file: $Version" -ForegroundColor Cyan
    } else {
        Write-Error "Version not specified and VERSION file not found at $VersionFile"
        exit 1
    }
}

function Copy-DirectoryFiles {
    param(
        [string]$SourceDir,
        [string]$DestinationDir
    )

    Get-ChildItem -Path $SourceDir -Recurse -File | ForEach-Object {
        $relativePath = $_.FullName.Substring($SourceDir.Length).TrimStart('\')
        $destinationPath = Join-Path $DestinationDir $relativePath
        $destinationParent = Split-Path -Parent $destinationPath

        if (-not (Test-Path $destinationParent)) {
            New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
        }

        Copy-Item -Path $_.FullName -Destination $destinationPath -Force
    }
}

function Apply-PluginOverlay {
    param(
        [string]$PluginName,
        [string]$DestinationDir
    )

    $overlaySource = Join-Path $PluginOverlayDir $PluginName
    if (-not (Test-Path $overlaySource)) {
        return
    }

    Write-Host "  Applying source-owned plugin overlay..." -ForegroundColor Cyan
    Copy-DirectoryFiles -SourceDir $overlaySource -DestinationDir $DestinationDir
}

Write-Host "`n=== Building Copilot CLI Plugins v$Version ===" -ForegroundColor Green
Write-Host "Source:   $RepoRoot"
Write-Host "Template: $PluginTemplateDir"
Write-Host "Output:   $OutputDir`n"

# Clean output
if (Test-Path $OutputDir) {
    Write-Host "Cleaning output: $OutputDir" -ForegroundColor Yellow
    Remove-Item -Path $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# =============================================================================
# Build: excel-mcp Plugin
# =============================================================================

Write-Host "`n[1/2] Building excel-mcp plugin..." -ForegroundColor Yellow

$TemplateMcp = Join-Path $PluginTemplateDir "excel-mcp"
$OutputMcp = Join-Path $OutputDir "excel-mcp"

if (-not (Test-Path $TemplateMcp)) {
    Write-Error "Template not found: $TemplateMcp"
    exit 1
}

Write-Host "  Copying validated plugin structure..." -ForegroundColor Cyan
Copy-Item -Path $TemplateMcp -Destination $OutputMcp -Recurse -Force

Apply-PluginOverlay -PluginName "excel-mcp" -DestinationDir $OutputMcp

Write-Host "  Updating plugin.json version to $Version..." -ForegroundColor Cyan
$PluginJsonPath = Join-Path $OutputMcp "plugin.json"
$pluginJson = Get-Content $PluginJsonPath -Raw | ConvertFrom-Json
$pluginJson.version = $Version
$pluginJson | ConvertTo-Json -Depth 10 | Set-Content $PluginJsonPath -Encoding UTF8

Write-Host "  Updating version.txt to $Version..." -ForegroundColor Cyan
Set-Content -Path (Join-Path $OutputMcp "version.txt") -Value $Version -Encoding UTF8 -NoNewline

Write-Host "  Refreshing excel-mcp skill content..." -ForegroundColor Cyan
$SourceSkillMcp = Join-Path $SkillsDir "excel-mcp\SKILL.md"
$DestSkillMcp = Join-Path $OutputMcp "skills\excel-mcp\SKILL.md"
Copy-Item -Path $SourceSkillMcp -Destination $DestSkillMcp -Force

Write-Host "  Refreshing shared references..." -ForegroundColor Cyan
$RefsDir = Join-Path $OutputMcp "skills\excel-mcp\references"
if (-not (Test-Path $RefsDir)) {
    New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
}
$SharedFiles = Get-ChildItem -Path $SharedDir -Filter "*.md"
foreach ($file in $SharedFiles) {
    Copy-Item -Path $file.FullName -Destination (Join-Path $RefsDir $file.Name) -Force
    Write-Host "    ✓ $($file.Name)" -ForegroundColor DarkGray
}

Write-Host "✅ excel-mcp plugin built" -ForegroundColor Green

# =============================================================================
# Build: excel-cli Plugin
# =============================================================================

Write-Host "`n[2/2] Building excel-cli plugin..." -ForegroundColor Yellow

$TemplateCli = Join-Path $PluginTemplateDir "excel-cli"
$OutputCli = Join-Path $OutputDir "excel-cli"

if (-not (Test-Path $TemplateCli)) {
    Write-Error "Template not found: $TemplateCli"
    exit 1
}

Write-Host "  Copying validated plugin structure..." -ForegroundColor Cyan
Copy-Item -Path $TemplateCli -Destination $OutputCli -Recurse -Force

Apply-PluginOverlay -PluginName "excel-cli" -DestinationDir $OutputCli

Write-Host "  Updating plugin.json version to $Version..." -ForegroundColor Cyan
$PluginJsonPath = Join-Path $OutputCli "plugin.json"
$pluginJson = Get-Content $PluginJsonPath -Raw | ConvertFrom-Json
$pluginJson.version = $Version
$pluginJson | ConvertTo-Json -Depth 10 | Set-Content $PluginJsonPath -Encoding UTF8

Write-Host "  Updating version.txt to $Version..." -ForegroundColor Cyan
Set-Content -Path (Join-Path $OutputCli "version.txt") -Value $Version -Encoding UTF8 -NoNewline

Write-Host "  Refreshing excel-cli skill content..." -ForegroundColor Cyan
$SourceSkillCli = Join-Path $SkillsDir "excel-cli\SKILL.md"
$DestSkillCli = Join-Path $OutputCli "skills\excel-cli\SKILL.md"
Copy-Item -Path $SourceSkillCli -Destination $DestSkillCli -Force

Write-Host "  Refreshing shared references..." -ForegroundColor Cyan
$RefsDir = Join-Path $OutputCli "skills\excel-cli\references"
if (-not (Test-Path $RefsDir)) {
    New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
}
$SharedFiles = Get-ChildItem -Path $SharedDir -Filter "*.md"
foreach ($file in $SharedFiles) {
    Copy-Item -Path $file.FullName -Destination (Join-Path $RefsDir $file.Name) -Force
    Write-Host "    ✓ $($file.Name)" -ForegroundColor DarkGray
}

Write-Host "✅ excel-cli plugin built" -ForegroundColor Green

# =============================================================================
# Summary
# =============================================================================

Write-Host "`n=== Build Complete ===" -ForegroundColor Green
Write-Host "Version: $Version"
Write-Host "Output:  $OutputDir"
Write-Host ""
Write-Host "Plugins:" -ForegroundColor Cyan
Write-Host "  ✓ excel-mcp (MCP Server + Skill)" -ForegroundColor Green
Write-Host "  ✓ excel-cli (CLI Skill only; install excelcli separately)" -ForegroundColor Green
Write-Host ""
Write-Host "Test locally:" -ForegroundColor Yellow
Write-Host "  copilot plugin install $OutputDir\excel-mcp"
Write-Host "  copilot plugin install $OutputDir\excel-cli"

<#
.SYNOPSIS
    Synchronizes the published Copilot plugin repo to the canonical marketplace layout.

.DESCRIPTION
    Copies built plugin artifacts into the published marketplace repo, applies any
    source-owned root overlay content, writes the canonical marketplace manifest to
    .github/plugin/marketplace.json, and removes the legacy root marketplace.json.

    The published repo is wrapper/bootstrap-only. Self-contained Windows runtimes
    remain in the main repo GitHub Releases and are acquired by plugin-local
    bootstrap logic on first invocation.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$PublishedRepoDir,

    [Parameter(Mandatory = $true)]
    [string]$BuiltPluginsDir,

    [Parameter(Mandatory = $true)]
    [string]$Version
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$RootOverlayDir = Join-Path $RepoRoot ".github\plugins\marketplace-repo"
$PublishedRepoDir = (Resolve-Path $PublishedRepoDir).Path
$BuiltPluginsDir = (Resolve-Path $BuiltPluginsDir).Path

function Copy-DirectoryFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,

        [Parameter(Mandatory = $true)]
        [string]$DestinationDir
    )

    Get-ChildItem -Path $SourceDir -Recurse -File -Force | ForEach-Object {
        $relativePath = $_.FullName.Substring($SourceDir.Length).TrimStart('\', '/')
        $destinationPath = Join-Path $DestinationDir $relativePath
        $destinationParent = Split-Path -Parent $destinationPath

        if (-not (Test-Path $destinationParent)) {
            New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
        }

        Copy-Item -Path $_.FullName -Destination $destinationPath -Force
    }
}

function Get-PluginSkillPaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PluginRoot,

        [Parameter(Mandatory = $true)]
        [string]$PluginName
    )

    $skillRoot = Join-Path $PluginRoot "skills"
    if (-not (Test-Path $skillRoot)) {
        return @()
    }

    return @(Get-ChildItem -Path $skillRoot -Directory |
        Where-Object { Test-Path (Join-Path $_.FullName "SKILL.md") } |
        Sort-Object Name |
        ForEach-Object { "./plugins/$PluginName/skills/$($_.Name)" })
}

function Write-Utf8NoBomJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        $Object
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $json = $Object | ConvertTo-Json -Depth 10
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, "$json`n", $utf8NoBom)
}

Write-Host "Synchronizing published plugin repo..." -ForegroundColor Cyan
Write-Host "  Published repo: $PublishedRepoDir" -ForegroundColor DarkGray
Write-Host "  Built plugins:  $BuiltPluginsDir" -ForegroundColor DarkGray
Write-Host "  Version:        $Version" -ForegroundColor DarkGray

$builtPluginNames = @("excel-mcp", "excel-cli")
$pluginMetadata = @()

foreach ($pluginName in $builtPluginNames) {
    $sourcePluginDir = Join-Path $BuiltPluginsDir $pluginName
    $pluginJsonPath = Join-Path $sourcePluginDir "plugin.json"

    if (-not (Test-Path $pluginJsonPath)) {
        throw "Built plugin manifest not found: $pluginJsonPath"
    }

    $pluginJson = Get-Content $pluginJsonPath -Raw | ConvertFrom-Json
    if ($pluginJson.version -ne $Version) {
        throw "$pluginJsonPath resolved version '$($pluginJson.version)' but expected '$Version'."
    }

    $versionTxtPath = Join-Path $sourcePluginDir "version.txt"
    if (Test-Path $versionTxtPath) {
        $resolvedVersionTxt = (Get-Content $versionTxtPath -Raw).Trim()
        if ($resolvedVersionTxt -ne $Version) {
            throw "$versionTxtPath resolved version '$resolvedVersionTxt' but expected '$Version'."
        }
    }

    $pluginMetadata += [ordered]@{
        name = $pluginJson.name
        source = "./plugins/$pluginName"
        description = $pluginJson.description
        version = $Version
        skills = @(Get-PluginSkillPaths -PluginRoot $sourcePluginDir -PluginName $pluginName)
    }
}

if (Test-Path $RootOverlayDir) {
    Write-Host "Applying source-owned published-repo overlay..." -ForegroundColor Cyan
    Copy-DirectoryFiles -SourceDir $RootOverlayDir -DestinationDir $PublishedRepoDir
}

foreach ($pluginName in $builtPluginNames) {
    $sourcePluginDir = Join-Path $BuiltPluginsDir $pluginName
    $destinationPluginDir = Join-Path $PublishedRepoDir "plugins\$pluginName"

    if (Test-Path $destinationPluginDir) {
        Remove-Item -Path $destinationPluginDir -Recurse -Force
    }

    Copy-Item -Path $sourcePluginDir -Destination $destinationPluginDir -Recurse -Force
}

$canonicalManifestPath = Join-Path $PublishedRepoDir ".github\plugin\marketplace.json"
$legacyManifestPath = Join-Path $PublishedRepoDir "marketplace.json"

$canonicalManifest = [ordered]@{
    name = "mcp-server-excel-plugins"
    metadata = [ordered]@{
        description = "Windows-only GitHub Copilot CLI plugins for Excel automation with ExcelMcp."
        version = "1.0.0"
    }
    owner = [ordered]@{
        name = "Stefan Brönner"
        email = "3026464+sbroenne@users.noreply.github.com"
    }
    plugins = $pluginMetadata
}

Write-Host "Writing canonical marketplace manifest..." -ForegroundColor Cyan
Write-Utf8NoBomJson -Path $canonicalManifestPath -Object $canonicalManifest

if (Test-Path $legacyManifestPath) {
    Write-Host "Removing legacy root marketplace manifest..." -ForegroundColor Cyan
    Remove-Item -Path $legacyManifestPath -Force
}

Write-Host "Published plugin repo synchronization complete." -ForegroundColor Green

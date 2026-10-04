<#
.SYNOPSIS
    Synchronizes the published Copilot plugin repo to the canonical marketplace layout.

.DESCRIPTION
    Copies built plugin artifacts into the published marketplace repo, applies any
    source-owned root overlay content, writes the canonical marketplace manifest to
    .github/plugin/marketplace.json, and removes the legacy root marketplace.json.

    The published repo contains npx launch configuration, an argument-safe CLI
    wrapper, and skills. Self-contained Windows runtimes remain outside the plugin.
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
$AgentPluginSchema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json"
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
Assert-PackageOutputPath -Path $PublishedRepoDir -RepoRoot $RepoRoot -Inputs @($BuiltPluginsDir)
if (-not (Test-Path -LiteralPath $RootOverlayDir -PathType Container)) { throw "Published-repository overlay is missing: $RootOverlayDir" }
foreach ($tree in @($PublishedRepoDir, $BuiltPluginsDir, $RootOverlayDir)) {
    if (((Get-Item -LiteralPath $tree -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -or
        @(Get-ChildItem -LiteralPath $tree -Recurse -Force | Where-Object {
        $_.Attributes -band [IO.FileAttributes]::ReparsePoint
    }).Count) { throw "Publication paths must not contain links: $tree" }
}

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

function Assert-AgentPluginManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PluginName,

        [Parameter(Mandatory = $true)]
        [string]$PluginJsonPath,

        [Parameter(Mandatory = $true)]
        $PluginJson
    )

    $allowedFields = @('$schema', "name", "version", "description", "author", "homepage", "repository", "license", "keywords", "extensions")
    foreach ($property in $PluginJson.PSObject.Properties) {
        if ($property.Name -notin $allowedFields) {
            throw "$PluginJsonPath contains unsupported Agent Plugins 1.0 field '$($property.Name)'."
        }
    }

    if ($PluginJson.'$schema' -ne $AgentPluginSchema) {
        throw "$PluginJsonPath must target $AgentPluginSchema."
    }
    if ($PluginJson.name -ne $PluginName) {
        throw "$PluginJsonPath has name '$($PluginJson.name)' but expected '$PluginName'."
    }
    if ($PluginJson.repository -isnot [string]) {
        throw "$PluginJsonPath repository must be a string."
    }
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
    Assert-AgentPluginManifest -PluginName $pluginName -PluginJsonPath $pluginJsonPath -PluginJson $pluginJson
    if ($pluginJson.version -ne $Version) {
        throw "$pluginJsonPath resolved version '$($pluginJson.version)' but expected '$Version'."
    }

    $requiredFiles = @(
        'README.md',
        'version.txt',
        "skills\$pluginName-report-formatting\SKILL.md",
        "skills\$pluginName-report-formatting\VERSION",
        "skills\$pluginName-report-formatting\references\report-formatting.md"
    )
    if ($pluginName -eq 'excel-cli') {
        $requiredFiles += 'bin\start-cli.ps1'
    } else {
        $requiredFiles += 'mcp.json'
    }
    foreach ($required in $requiredFiles) {
        if (-not (Test-Path -LiteralPath (Join-Path $sourcePluginDir $required) -PathType Leaf)) {
            throw "Incomplete plugin payload: $pluginName is missing $required."
        }
    }
    $skillVersion = (Get-Content -LiteralPath (Join-Path $sourcePluginDir "skills\$pluginName-report-formatting\VERSION") -Raw).Trim()
    if ($skillVersion -ne $Version) { throw "Prepared $pluginName skill version must match $Version." }

    $globalHelpers = @(Get-ChildItem -LiteralPath $sourcePluginDir -Recurse -Force -File -Filter "install-global.ps1")
    if ($globalHelpers.Count) {
        throw "Global installation helpers are retired; use npx instead: $($globalHelpers.FullName -join ', ')"
    }

    $legacyMcpPath = Join-Path $sourcePluginDir ".mcp.json"
    if (Test-Path $legacyMcpPath) {
        throw "Legacy MCP configuration is not permitted in Agent Plugins 1.0 packages: $legacyMcpPath"
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
        author = $pluginJson.author
        homepage = $pluginJson.homepage
        repository = $pluginJson.repository
        license = $pluginJson.license
        keywords = @($pluginJson.keywords)
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

    New-Item -ItemType Directory -Path (Split-Path $destinationPluginDir -Parent) -Force | Out-Null
    Install-PackageOutput -Source $sourcePluginDir -Destination $destinationPluginDir
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

# Claude Code reads .claude-plugin/marketplace.json. Plugin directories carry no
# .claude-plugin/plugin.json, so each entry is the manifest and declares the MCP server inline.
$claudeEntries = @(foreach ($entry in $pluginMetadata) {
    $claudeEntry = [ordered]@{
        name = $entry.name
        source = $entry.source
        description = $entry.description
        version = $entry.version
        author = $entry.author
        homepage = $entry.homepage
        repository = $entry.repository
        license = $entry.license
        keywords = $entry.keywords
    }
    $mcpConfigPath = Join-Path $BuiltPluginsDir "$($entry.name)\mcp.json"
    if (Test-Path -LiteralPath $mcpConfigPath -PathType Leaf) {
        $claudeEntry.mcpServers = (Get-Content -LiteralPath $mcpConfigPath -Raw | ConvertFrom-Json).mcpServers
    }
    $claudeEntry
})
$claudeManifest = [ordered]@{
    name = "mcp-server-excel-plugins"
    description = "Windows-only Claude Code plugins for Excel automation with ExcelMcp."
    owner = $canonicalManifest.owner
    plugins = $claudeEntries
}

Write-Host "Writing Claude Code marketplace manifest..." -ForegroundColor Cyan
Write-Utf8NoBomJson -Path (Join-Path $PublishedRepoDir ".claude-plugin\marketplace.json") -Object $claudeManifest

if (Test-Path $legacyManifestPath) {
    Write-Host "Removing legacy root marketplace manifest..." -ForegroundColor Cyan
    Remove-Item -Path $legacyManifestPath -Force
}

Write-Host "Published plugin repo synchronization complete." -ForegroundColor Green

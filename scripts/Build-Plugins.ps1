<#
.SYNOPSIS
    Builds Agent Plugins from canonical source-owned templates and updates versions.

.DESCRIPTION
    1. Copy canonical plugin templates from .github/plugins/
    2. Strip any runtime payloads from plugin bin/ roots
    3. Update runtime-bootstrap metadata in plugin.json and version.txt
    4. Synchronize complete Agent Skill directories from source
    5. Validate Agent Plugins 1.0 and Agent Skills layout requirements

    RUNTIME BOOTSTRAP MODEL:
    - Published plugins ship wrapper/download logic and metadata only
    - Self-contained Windows runtimes are downloaded from the latest GitHub release on first use
    - No committed .exe/.dll runtime payloads should survive into the published plugin repo

    OUTPUT:
    plugins/
      excel-mcp/     → MCP plugin (wrapper/bootstrap assets + updated version + fresh skills)
      excel-cli/     → CLI plugin (wrapper/bootstrap assets + updated version + fresh skills)

.PARAMETER Version
    Plugin version. Required for distributable builds.

.PARAMETER OutputDir
    Output directory. Default: plugins/

.EXAMPLE
    ./Build-Plugins.ps1 -Version 1.2.3
#>
param(
    [string]$Version = $null,
    [string]$OutputDir = "plugins"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$PluginSourceDir = Join-Path $RepoRoot ".github\plugins"
$AgentPluginSchema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json"
$AgentPluginMcpSchema = "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json"

if ([string]::IsNullOrWhiteSpace($Version)) {
    throw "Version is required. Pass -Version <version>."
}
$Version = $Version.Trim()

$BootstrapScriptPath = Join-Path $RepoRoot "scripts\Build-BootstrapScripts.ps1"
if (-not (Test-Path $BootstrapScriptPath)) {
    throw "Bootstrap generator script not found: $BootstrapScriptPath"
}

Write-Host "Rendering canonical plugin bootstrap scripts from the shared template..." -ForegroundColor Cyan
& $BootstrapScriptPath -OutputRoot $PluginSourceDir
if ($LASTEXITCODE -ne 0) {
    throw "Bootstrap script generation failed."
}

function Remove-PackagedRuntimePayload {
    param(
        [string]$PluginName,
        [string]$PluginDir
    )

    $pluginBinDir = Join-Path $PluginDir "bin"
    if (-not (Test-Path $pluginBinDir)) {
        return
    }

    $runtimePayload = Get-ChildItem -Path $pluginBinDir -Recurse -Force -File | Where-Object {
        $_.Extension -in @(".exe", ".dll", ".pdb") -or
        $_.Name.EndsWith(".deps.json", [System.StringComparison]::OrdinalIgnoreCase) -or
        $_.Name.EndsWith(".runtimeconfig.json", [System.StringComparison]::OrdinalIgnoreCase)
    }

    foreach ($file in $runtimePayload) {
        Write-Host "  Removing committed runtime payload from ${PluginName}: $($file.FullName)" -ForegroundColor DarkYellow
        Remove-Item -Path $file.FullName -Force
    }
}

function Update-PluginManifest {
    param(
        [string]$PluginJsonPath,
        [string]$Version
    )

    $pluginJson = Get-Content $PluginJsonPath -Raw | ConvertFrom-Json
    $pluginJson.version = $Version
    $pluginJson | ConvertTo-Json -Depth 10 | Set-Content $PluginJsonPath -Encoding UTF8
}

function Copy-AgentSkill {
    param(
        [string]$SourceDir,
        [string]$DestinationDir,
        [string]$Version = $null
    )

    if (-not (Test-Path $SourceDir -PathType Container)) {
        throw "Canonical Agent Skill directory not found: $SourceDir"
    }

    if (Test-Path $DestinationDir) {
        Remove-Item -Path $DestinationDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path (Split-Path -Parent $DestinationDir) -Force | Out-Null
    Copy-Item -Path $SourceDir -Destination $DestinationDir -Recurse -Force

    # Write the VERSION file unconditionally. Guarding on Test-Path only *updated* a VERSION that
    # the canonical skill already carried, so excel-cli - whose source skill has none - shipped
    # without one while excel-mcp shipped with one.
    $skillVersionPath = Join-Path $DestinationDir "VERSION"
    if ($Version) {
        Set-Content -Path $skillVersionPath -Value $Version -Encoding UTF8 -NoNewline
    }
}

function Assert-AgentSkill {
    param(
        [string]$SkillDir,
        [string]$ExpectedVersion
    )

    $skillPath = Join-Path $SkillDir "SKILL.md"
    if (-not (Test-Path $skillPath -PathType Leaf)) {
        throw "Agent Skill is missing SKILL.md: $SkillDir"
    }

    # Every packaged skill must carry a VERSION stamped with the version the plugin was built at.
    # excel-cli previously shipped with no VERSION at all, and nothing failed the build.
    if ($ExpectedVersion) {
        $versionPath = Join-Path $SkillDir "VERSION"
        if (-not (Test-Path $versionPath -PathType Leaf)) {
            throw "Agent Skill is missing VERSION: $SkillDir"
        }

        $skillVersion = (Get-Content $versionPath -Raw).Trim()
        if ($skillVersion -ne $ExpectedVersion) {
            throw "$versionPath has version '$skillVersion' but expected '$ExpectedVersion'."
        }
    }

    $lines = @(Get-Content $skillPath)
    if ($lines.Count -lt 3 -or $lines[0].Trim() -ne "---") {
        throw "$skillPath must begin with YAML frontmatter."
    }

    $closingDelimiter = -1
    for ($index = 1; $index -lt $lines.Count; $index++) {
        if ($lines[$index].Trim() -eq "---") {
            $closingDelimiter = $index
            break
        }
    }

    if ($closingDelimiter -lt 2) {
        throw "$skillPath is missing the closing YAML frontmatter delimiter."
    }

    $frontmatter = $lines[1..($closingDelimiter - 1)]
    $allowedFields = @("name", "description", "license", "compatibility", "metadata", "allowed-tools")
    $fields = @{}
    foreach ($line in $frontmatter) {
        if ($line -match "^([a-z][a-z-]*):(?:\s*(.*))?$") {
            $fieldName = $Matches[1]
            if ($fieldName -notin $allowedFields) {
                throw "$skillPath contains unsupported Agent Skills frontmatter field '$fieldName'."
            }
            $fields[$fieldName] = $Matches[2]
        }
    }

    $expectedName = Split-Path -Leaf $SkillDir
    $skillName = $fields["name"]
    if ($skillName -ne $expectedName -or $skillName -notmatch "^(?!.*--)[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$" -or $skillName.Length -gt 64) {
        throw "$skillPath has invalid Agent Skill name '$skillName'; it must match directory '$expectedName'."
    }

    if (-not $fields.ContainsKey("description")) {
        throw "$skillPath must declare an Agent Skill description."
    }

    $descriptionStart = [Array]::IndexOf($frontmatter, ($frontmatter | Where-Object { $_ -match "^description:" } | Select-Object -First 1))
    $description = $fields["description"]
    if ($description -in @(">", "|")) {
        $descriptionLines = @()
        for ($index = $descriptionStart + 1; $index -lt $frontmatter.Count; $index++) {
            if ($frontmatter[$index] -notmatch "^\s+") {
                break
            }
            $descriptionLines += $frontmatter[$index].Trim()
        }
        $description = $descriptionLines -join " "
    }

    if ([string]::IsNullOrWhiteSpace($description) -or $description.Length -gt 1024) {
        throw "$skillPath has an invalid Agent Skill description length."
    }
}

function Assert-AgentPluginPackage {
    param(
        [string]$PluginName,
        [string]$PluginDir,
        [string]$ExpectedVersion
    )

    $pluginJsonPath = Join-Path $PluginDir "plugin.json"
    if (-not (Test-Path $pluginJsonPath -PathType Leaf)) {
        throw "Agent Plugin manifest not found: $pluginJsonPath"
    }

    $pluginJson = Get-Content $pluginJsonPath -Raw | ConvertFrom-Json
    $allowedFields = @('$schema', "name", "version", "description", "author", "homepage", "repository", "license", "keywords", "extensions")
    foreach ($property in $pluginJson.PSObject.Properties) {
        if ($property.Name -notin $allowedFields) {
            throw "$pluginJsonPath contains unsupported Agent Plugins 1.0 field '$($property.Name)'."
        }
    }

    if ($pluginJson.'$schema' -ne $AgentPluginSchema) {
        throw "$pluginJsonPath must target $AgentPluginSchema."
    }
    if ($pluginJson.name -ne $PluginName) {
        throw "$pluginJsonPath has name '$($pluginJson.name)' but expected '$PluginName'."
    }
    if ($pluginJson.version -ne $ExpectedVersion) {
        throw "$pluginJsonPath has version '$($pluginJson.version)' but expected '$ExpectedVersion'."
    }
    if ($pluginJson.repository -isnot [string]) {
        throw "$pluginJsonPath repository must be a string."
    }

    $legacyCopilotHelper = Join-Path $PluginDir "bin\install-global.ps1"
    if (Test-Path $legacyCopilotHelper) {
        throw "Copilot-only files must be placed under com.github.copilot/: $legacyCopilotHelper"
    }

    $legacyMcpPath = Join-Path $PluginDir ".mcp.json"
    if (Test-Path $legacyMcpPath) {
        throw "Legacy MCP configuration is not permitted in Agent Plugins 1.0 packages: $legacyMcpPath"
    }

    $skillsRoot = Join-Path $PluginDir "skills"
    if (Test-Path $skillsRoot) {
        Get-ChildItem -Path $skillsRoot -Directory | ForEach-Object {
            Assert-AgentSkill -SkillDir $_.FullName -ExpectedVersion $ExpectedVersion
        }
    }

    $mcpPath = Join-Path $PluginDir "mcp.json"
    if (-not (Test-Path $mcpPath)) {
        return
    }

    $mcp = Get-Content $mcpPath -Raw | ConvertFrom-Json
    $mcpProperties = @($mcp.PSObject.Properties.Name)
    if ($mcpProperties.Count -ne 2 -or '$schema' -notin $mcpProperties -or "mcpServers" -notin $mcpProperties) {
        throw "$mcpPath must contain only '`$schema' and 'mcpServers'."
    }
    if ($mcp.'$schema' -ne $AgentPluginMcpSchema) {
        throw "$mcpPath must target $AgentPluginMcpSchema."
    }

    foreach ($serverProperty in $mcp.mcpServers.PSObject.Properties) {
        $server = $serverProperty.Value
        $serverFields = @($server.PSObject.Properties.Name)
        $allowedServerFields = @("type", "command", "args", "env", "cwd")
        if ($serverFields | Where-Object { $_ -notin $allowedServerFields }) {
            throw "$mcpPath server '$($serverProperty.Name)' contains unsupported fields."
        }
        if ($server.type -ne "stdio" -or $server.command -isnot [string] -or $server.command -match "\s") {
            throw "$mcpPath server '$($serverProperty.Name)' must use type 'stdio' and a single executable command token."
        }
        if (@($server.args) | Where-Object { $_ -match "\{pluginDir\}" }) {
            throw "$mcpPath server '$($serverProperty.Name)' still uses the legacy '{pluginDir}' placeholder."
        }
    }
}

Write-Host "`n=== Building Agent Plugins v$Version ===" -ForegroundColor Green
Write-Host "Source:   $RepoRoot"
Write-Host "Templates: $PluginSourceDir"
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

$TemplateMcp = Join-Path $PluginSourceDir "excel-mcp"
$OutputMcp = Join-Path $OutputDir "excel-mcp"

if (-not (Test-Path $TemplateMcp)) {
    Write-Error "Template not found: $TemplateMcp"
    exit 1
}

Write-Host "  Copying canonical plugin template..." -ForegroundColor Cyan
Copy-Item -Path $TemplateMcp -Destination $OutputMcp -Recurse -Force

Remove-PackagedRuntimePayload -PluginName "excel-mcp" -PluginDir $OutputMcp

Write-Host "  Updating plugin.json version to $Version..." -ForegroundColor Cyan
$PluginJsonPath = Join-Path $OutputMcp "plugin.json"
Update-PluginManifest `
    -PluginJsonPath $PluginJsonPath `
    -Version $Version

Write-Host "  Updating version.txt to $Version..." -ForegroundColor Cyan
Set-Content -Path (Join-Path $OutputMcp "version.txt") -Value $Version -Encoding UTF8 -NoNewline

Write-Host "  Synchronizing complete excel-mcp skill directory..." -ForegroundColor Cyan
$SourceSkillMcp = Join-Path $SkillsDir "excel-mcp"
$DestSkillMcp = Join-Path $OutputMcp "skills\excel-mcp"
Copy-AgentSkill -SourceDir $SourceSkillMcp -DestinationDir $DestSkillMcp -Version $Version

Assert-AgentPluginPackage -PluginName "excel-mcp" -PluginDir $OutputMcp -ExpectedVersion $Version
Write-Host "✅ excel-mcp plugin built" -ForegroundColor Green

# =============================================================================
# Build: excel-cli Plugin
# =============================================================================

Write-Host "`n[2/2] Building excel-cli plugin..." -ForegroundColor Yellow

$TemplateCli = Join-Path $PluginSourceDir "excel-cli"
$OutputCli = Join-Path $OutputDir "excel-cli"

if (-not (Test-Path $TemplateCli)) {
    Write-Error "Template not found: $TemplateCli"
    exit 1
}

Write-Host "  Copying canonical plugin template..." -ForegroundColor Cyan
Copy-Item -Path $TemplateCli -Destination $OutputCli -Recurse -Force

Remove-PackagedRuntimePayload -PluginName "excel-cli" -PluginDir $OutputCli

Write-Host "  Updating plugin.json version to $Version..." -ForegroundColor Cyan
$PluginJsonPath = Join-Path $OutputCli "plugin.json"
Update-PluginManifest `
    -PluginJsonPath $PluginJsonPath `
    -Version $Version

Write-Host "  Updating version.txt to $Version..." -ForegroundColor Cyan
Set-Content -Path (Join-Path $OutputCli "version.txt") -Value $Version -Encoding UTF8 -NoNewline

Write-Host "  Synchronizing complete excel-cli skill directory..." -ForegroundColor Cyan
$SourceSkillCli = Join-Path $SkillsDir "excel-cli"
$DestSkillCli = Join-Path $OutputCli "skills\excel-cli"
Copy-AgentSkill -SourceDir $SourceSkillCli -DestinationDir $DestSkillCli -Version $Version

Assert-AgentPluginPackage -PluginName "excel-cli" -PluginDir $OutputCli -ExpectedVersion $Version
Write-Host "✅ excel-cli plugin built" -ForegroundColor Green

# =============================================================================
# Summary
# =============================================================================

Write-Host "`n=== Build Complete ===" -ForegroundColor Green
Write-Host "Version: $Version"
Write-Host "Output:  $OutputDir"
Write-Host ""
Write-Host "Plugins:" -ForegroundColor Cyan
Write-Host '  [ok] excel-mcp - bootstrap assets and skill' -ForegroundColor Green
Write-Host '  [ok] excel-cli - bootstrap assets and skill' -ForegroundColor Green
Write-Host ""
Write-Host "Test locally:" -ForegroundColor Yellow
Write-Host "  copilot plugin install $OutputDir\excel-mcp"
Write-Host "  copilot plugin install $OutputDir\excel-cli"

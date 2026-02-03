<#
.SYNOPSIS
    Builds the Excel MCP Agent Skills packages for distribution.

.DESCRIPTION
    Creates distributable artifacts for Agent Skills:
    - excel-mcp-skill-v{version}.zip: MCP Server skill package
    - excel-cli-skill-v{version}.zip: CLI skill package
    - CLAUDE.md: Claude Code project instructions
    - .cursorrules: Cursor project rules

    Shared behavioral guidance from skills/shared/ is automatically copied
    to both excel-mcp/references/ and excel-cli/references/ during packaging.

.PARAMETER OutputDir
    Output directory for artifacts. Default: artifacts/skills

.PARAMETER Version
    Override version from skills/excel-mcp/VERSION

.PARAMETER PopulateReferences
    Copy shared references to skill folders for local development (without packaging).

.EXAMPLE
    ./Build-AgentSkills.ps1

.EXAMPLE
    ./Build-AgentSkills.ps1 -OutputDir ./dist -Version 1.2.0

.EXAMPLE
    ./Build-AgentSkills.ps1 -PopulateReferences
#>
param(
    [string]$OutputDir = "artifacts/skills",
    [string]$Version = $null,
    [switch]$PopulateReferences
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$SharedDir = Join-Path $SkillsDir "shared"

# Function to copy shared references to a skill's references folder
function Copy-SharedReferences {
    param(
        [string]$SkillPath,
        [string]$SkillName
    )

    $RefsDir = Join-Path $SkillPath "references"

    # Create references directory if it doesn't exist
    if (-not (Test-Path $RefsDir)) {
        New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
    }

    # Define which files each skill needs (based on SKILL.md @references/)
    $SkillReferences = @{
        "excel-cli" = @(
            "behavioral-rules.md"
            "anti-patterns.md"
            "workflows.md"
        )
        "excel-mcp" = @(
            "behavioral-rules.md"
            "anti-patterns.md"
            "workflows.md"
            "excel_chart.md"
            "excel_conditionalformat.md"
            "excel_datamodel.md"
            "excel_powerquery.md"
            "excel_range.md"
            "excel_slicer.md"
            "excel_table.md"
            "excel_worksheet.md"
        )
    }

    # Get the list of files for this skill
    $FilesToCopy = $SkillReferences[$SkillName]
    if (-not $FilesToCopy) {
        Write-Warning "No reference files defined for skill: $SkillName"
        return
    }

    # Copy only the files this skill needs
    if (Test-Path $SharedDir) {
        $CopiedCount = 0
        foreach ($fileName in $FilesToCopy) {
            $sourceFile = Join-Path $SharedDir $fileName
            if (Test-Path $sourceFile) {
                Copy-Item -Path $sourceFile -Destination $RefsDir -Force
                $CopiedCount++
            } else {
                Write-Warning "Reference file not found in shared: $fileName"
            }
        }
        Write-Host "  Copied $CopiedCount shared references to $SkillName/references/" -ForegroundColor Green
    } else {
        Write-Warning "Shared directory not found: $SharedDir"
    }
}

# Handle -PopulateReferences mode (for development)
if ($PopulateReferences) {
    Write-Host "Populating references from shared/ for local development..." -ForegroundColor Cyan

    # Copy to excel-mcp
    $McpPath = Join-Path $SkillsDir "excel-mcp"
    if (Test-Path $McpPath) {
        Copy-SharedReferences -SkillPath $McpPath -SkillName "excel-mcp"
    }

    # Copy to excel-cli
    $CliPath = Join-Path $SkillsDir "excel-cli"
    if (Test-Path $CliPath) {
        Copy-SharedReferences -SkillPath $CliPath -SkillName "excel-cli"
    }

    Write-Host ""
    Write-Host "Done! References populated for local development." -ForegroundColor Green
    exit 0
}

# Get version
if (-not $Version) {
    $VersionFile = Join-Path $SkillsDir "excel-mcp/VERSION"
    if (Test-Path $VersionFile) {
        $Version = (Get-Content $VersionFile -Raw).Trim()
    } else {
        $Version = "0.0.0"
    }
}

Write-Host "Building Agent Skills packages v$Version" -ForegroundColor Cyan
Write-Host "Source: $SkillsDir"
Write-Host "Output: $OutputDir"
Write-Host ""

# Create output directory
$OutputPath = Join-Path $RepoRoot $OutputDir
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Function to build a single skill package
function Build-SkillPackage {
    param(
        [string]$SkillName,
        [string]$SkillDescription
    )

    $SkillSource = Join-Path $SkillsDir $SkillName
    if (-not (Test-Path $SkillSource)) {
        Write-Warning "Skill not found: $SkillName"
        return
    }

    Write-Host "Building $SkillName package..." -ForegroundColor Yellow

    # Create staging directory for this skill
    $StagingDir = Join-Path $env:TEMP "$SkillName-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

    try {
        # Copy skill content
        Copy-Item -Path $SkillSource -Destination "$StagingDir/$SkillName" -Recurse
        Write-Host "  Copied skill content" -ForegroundColor Green

        # Copy shared references
        Copy-SharedReferences -SkillPath "$StagingDir/$SkillName" -SkillName $SkillName

        # Copy skill-specific README if it exists, otherwise use the skill's SKILL.md as basis
        $SkillReadme = Join-Path $SkillSource "README.md"
        if (Test-Path $SkillReadme) {
            Copy-Item -Path $SkillReadme -Destination $StagingDir
        }

        # Create ZIP archive
        $ZipName = "$SkillName-skill-v$Version.zip"
        $ZipPath = Join-Path $OutputPath $ZipName

        if (Test-Path $ZipPath) {
            Remove-Item $ZipPath -Force
        }

        Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal
        Write-Host "  Created: $ZipName" -ForegroundColor Green

        return $ZipName
    } finally {
        if (Test-Path $StagingDir) {
            Remove-Item $StagingDir -Recurse -Force
        }
    }
}

# Build each skill package
$McpZip = Build-SkillPackage -SkillName "excel-mcp" -SkillDescription "MCP Server skill"
$CliZip = Build-SkillPackage -SkillName "excel-cli" -SkillDescription "CLI skill"

# Copy CLAUDE.md and .cursorrules
Write-Host "Copying platform-specific files..." -ForegroundColor Yellow

$ClaudeSrc = Join-Path $SkillsDir "CLAUDE.md"
if (Test-Path $ClaudeSrc) {
    Copy-Item -Path $ClaudeSrc -Destination $OutputPath
    Write-Host "  Created: CLAUDE.md" -ForegroundColor Green
}

$CursorSrc = Join-Path $SkillsDir ".cursorrules"
if (Test-Path $CursorSrc) {
    Copy-Item -Path $CursorSrc -Destination $OutputPath
    Write-Host "  Created: .cursorrules" -ForegroundColor Green
}

# Generate manifest
$Manifest = @{
    name = "excel-mcp-skills"
    version = $Version
    description = "Excel MCP Server Agent Skills for AI coding assistants"
    platforms = @("github-copilot", "claude-code", "cursor", "windsurf", "gemini-cli", "goose")
    skills = @(
        @{
            name = "excel-mcp"
            file = $McpZip
            description = "MCP Server skill - for conversational AI (Claude Desktop, VS Code Chat)"
            target = "MCP Server"
        }
        @{
            name = "excel-cli"
            file = $CliZip
            description = "CLI skill - for coding agents (Copilot, Cursor, Windsurf)"
            target = "CLI Tool"
        }
    )
    files = @(
        @{ name = "CLAUDE.md"; type = "config"; description = "Claude Code project instructions" }
        @{ name = ".cursorrules"; type = "config"; description = "Cursor project rules" }
    )
    repository = "https://github.com/sbroenne/mcp-server-excel"
    documentation = "https://excelmcpserver.dev/"
    buildDate = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
}

$ManifestPath = Join-Path $OutputPath "manifest.json"
$Manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $ManifestPath -Encoding UTF8
Write-Host "  Created: manifest.json" -ForegroundColor Green

Write-Host ""
Write-Host "Build complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Output files in: $OutputPath" -ForegroundColor Cyan
Get-ChildItem $OutputPath | ForEach-Object {
    $Size = if ($_.Length -gt 1MB) { "{0:N2} MB" -f ($_.Length / 1MB) }
            elseif ($_.Length -gt 1KB) { "{0:N2} KB" -f ($_.Length / 1KB) }
            else { "{0} bytes" -f $_.Length }
    Write-Host "  $($_.Name) ($Size)"
}

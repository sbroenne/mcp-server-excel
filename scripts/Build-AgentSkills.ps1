<#
.SYNOPSIS
    Builds the Excel MCP Agent Skills package for distribution.

.DESCRIPTION
    Creates distributable artifacts for the Agent Skills package:
    - excel-mcp-skills.zip: Full skill package for manual installation
    - CLAUDE.md: Claude Code project instructions
    - .cursorrules: Cursor project rules

.PARAMETER OutputDir
    Output directory for artifacts. Default: artifacts/skills

.PARAMETER Version
    Override version from skills/excel-mcp/VERSION

.EXAMPLE
    ./Build-AgentSkills.ps1

.EXAMPLE
    ./Build-AgentSkills.ps1 -OutputDir ./dist -Version 1.2.0
#>
param(
    [string]$OutputDir = "artifacts/skills",
    [string]$Version = $null
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"

# Get version
if (-not $Version) {
    $VersionFile = Join-Path $SkillsDir "excel-mcp/VERSION"
    if (Test-Path $VersionFile) {
        $Version = (Get-Content $VersionFile -Raw).Trim()
    } else {
        $Version = "0.0.0"
    }
}

Write-Host "Building Agent Skills package v$Version" -ForegroundColor Cyan
Write-Host "Source: $SkillsDir"
Write-Host "Output: $OutputDir"
Write-Host ""

# Create output directory
$OutputPath = Join-Path $RepoRoot $OutputDir
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Create staging directory
$StagingDir = Join-Path $env:TEMP "excel-mcp-skills-$([guid]::NewGuid().ToString('N').Substring(0,8))"
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

try {
    # Copy skill content to staging
    Write-Host "Copying skill content..." -ForegroundColor Yellow

    $SkillSource = Join-Path $SkillsDir "excel-mcp"
    $SkillDest = Join-Path $StagingDir "excel-mcp"
    Copy-Item -Path $SkillSource -Destination $SkillDest -Recurse

    # Copy root-level files
    $ReadmeSrc = Join-Path $SkillsDir "README.md"
    if (Test-Path $ReadmeSrc) {
        Copy-Item -Path $ReadmeSrc -Destination $StagingDir
    }

    # Copy CLAUDE.md
    $ClaudeSrc = Join-Path $SkillsDir "CLAUDE.md"
    if (Test-Path $ClaudeSrc) {
        Copy-Item -Path $ClaudeSrc -Destination $StagingDir
        Copy-Item -Path $ClaudeSrc -Destination $OutputPath
        Write-Host "  Created: CLAUDE.md" -ForegroundColor Green
    }

    # Copy .cursorrules
    $CursorSrc = Join-Path $SkillsDir ".cursorrules"
    if (Test-Path $CursorSrc) {
        Copy-Item -Path $CursorSrc -Destination $StagingDir
        Copy-Item -Path $CursorSrc -Destination $OutputPath
        Write-Host "  Created: .cursorrules" -ForegroundColor Green
    }

    # Create ZIP archive
    Write-Host "Creating ZIP archive..." -ForegroundColor Yellow
    $ZipName = "excel-mcp-skills-v$Version.zip"
    $ZipPath = Join-Path $OutputPath $ZipName

    if (Test-Path $ZipPath) {
        Remove-Item $ZipPath -Force
    }

    Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal
    Write-Host "  Created: $ZipName" -ForegroundColor Green

    # Also create a latest symlink/copy
    $LatestZip = Join-Path $OutputPath "excel-mcp-skills.zip"
    Copy-Item -Path $ZipPath -Destination $LatestZip -Force
    Write-Host "  Created: excel-mcp-skills.zip (latest)" -ForegroundColor Green

    # Generate manifest
    $Manifest = @{
        name = "excel-mcp-skills"
        version = $Version
        description = "Excel MCP Server Agent Skills for AI coding assistants"
        platforms = @("github-copilot", "claude-code", "cursor", "windsurf", "gemini-cli", "goose")
        files = @(
            @{ name = $ZipName; type = "package"; description = "Full skill package" }
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

} finally {
    # Cleanup staging directory
    if (Test-Path $StagingDir) {
        Remove-Item $StagingDir -Recurse -Force
    }
}

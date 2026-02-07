<#
.SYNOPSIS
    Builds the Excel MCP Agent Skills package for distribution.

.DESCRIPTION
    Creates a single distributable artifact for Agent Skills:
    - excel-skills-v{version}.zip: Combined skill package with both excel-mcp and excel-cli
    - CLAUDE.md: Claude Code project instructions
    - .cursorrules: Cursor project rules

    Shared behavioral guidance from skills/shared/ is automatically copied
    to both excel-mcp/references/ and excel-cli/references/ during packaging.

    Users install with: npx skills add sbroenne/mcp-server-excel
    Then select which skill(s) they want: excel-cli, excel-mcp, or both.

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

# Function to generate CLI command reference from excelcli --help output
function Generate-CliReference {
    param(
        [string]$SkillPath,
        [string]$ExcelCliPath = $null
    )

    # Find excelcli binary
    if (-not $ExcelCliPath) {
        $ExcelCliPath = Join-Path $RepoRoot "src/ExcelMcp.CLI/bin/Release/net10.0-windows/excelcli.exe"
    }

    if (-not (Test-Path $ExcelCliPath)) {
        Write-Warning "excelcli not found at $ExcelCliPath - skipping CLI reference generation"
        Write-Warning "Build the CLI first: dotnet build src/ExcelMcp.CLI -c Release"
        return
    }

    Write-Host "  Generating CLI command reference from excelcli..." -ForegroundColor Cyan

    $RefsDir = Join-Path $SkillPath "references"
    if (-not (Test-Path $RefsDir)) {
        New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
    }

    $OutputFile = Join-Path $RefsDir "cli-commands.md"
    $Content = @()
    $Content += "# CLI Command Reference"
    $Content += ""
    $Content += "> Auto-generated from \`excelcli --help\`. Do not edit manually."
    $Content += ""

    # Get main help to extract commands
    $MainHelp = & $ExcelCliPath --help 2>&1 | Out-String

    # Parse commands from main help (look for lines with command names)
    $Commands = @()
    $InCommands = $false
    foreach ($line in ($MainHelp -split "`r?`n")) {
        if ($line -match "^COMMANDS:") {
            $InCommands = $true
            continue
        }
        if ($InCommands -and $line -match "^\s{4}(\w+)\s") {
            $Commands += $Matches[1]
        }
    }

    # Skip 'session' as it's documented separately
    $Commands = $Commands | Where-Object { $_ -ne "session" }

    foreach ($cmd in $Commands) {
        $Content += "## $cmd"
        $Content += ""

        # Get command help
        $CmdHelp = & $ExcelCliPath $cmd --help 2>&1 | Out-String

        # Extract actions from the description line
        if ($CmdHelp -match "Actions:\s*(.+?)(?:\r?\n|$)") {
            $ActionsStr = $Matches[1] -replace "\s+", " "
            $Actions = ($ActionsStr -split ",\s*") | ForEach-Object { $_.Trim().TrimEnd('.') } | Where-Object { $_ -ne "" }

            # Extract options section
            $Options = @()
            $InOptions = $false
            foreach ($line in ($CmdHelp -split "`r?`n")) {
                if ($line -match "^OPTIONS:") {
                    $InOptions = $true
                    continue
                }
                if ($InOptions -and $line -match "^\s+--(\S+)\s+<[^>]+>\s+(.+)$") {
                    $Options += @{
                        Name = $Matches[1]
                        Desc = $Matches[2].Trim()
                    }
                }
                elseif ($InOptions -and $line -match "^\s+-\w\|--(\S+)\s+<[^>]+>\s+(.+)$") {
                    $Options += @{
                        Name = $Matches[1]
                        Desc = $Matches[2].Trim()
                    }
                }
            }

            # Output actions
            $ActionsList = $Actions | ForEach-Object { "``$_``" }
            $Content += "**Actions:** $($ActionsList -join ', ')"
            $Content += ""

            # Output parameters table if any
            if ($Options.Count -gt 0) {
                $Content += "| Parameter | Description |"
                $Content += "|-----------|-------------|"
                foreach ($opt in $Options) {
                    $Content += "| ``--$($opt.Name)`` | $($opt.Desc) |"
                }
                $Content += ""
            }
        }
    }

    # Write the file
    $Content -join "`n" | Set-Content -Path $OutputFile -Encoding UTF8 -NoNewline
    Write-Host "  Generated: cli-commands.md" -ForegroundColor Green
}

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
            # cli-commands.md is generated dynamically by Generate-CliReference
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
        # Generate CLI command reference from excelcli --help
        Generate-CliReference -SkillPath $CliPath
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

Write-Host "Building Agent Skills package v$Version" -ForegroundColor Cyan
Write-Host "Source: $SkillsDir"
Write-Host "Output: $OutputDir"
Write-Host ""

# Create output directory
$OutputPath = Join-Path $RepoRoot $OutputDir
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Build combined skills package
Write-Host "Building combined skills package..." -ForegroundColor Yellow

# Create staging directory
$StagingDir = Join-Path $env:TEMP "excel-skills-$([guid]::NewGuid().ToString('N').Substring(0,8))"
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

try {
    # Create skills/ subdirectory (the standard location for npx skills add)
    $SkillsStagingDir = Join-Path $StagingDir "skills"
    New-Item -ItemType Directory -Path $SkillsStagingDir -Force | Out-Null

    # Copy excel-mcp skill
    $McpSource = Join-Path $SkillsDir "excel-mcp"
    if (Test-Path $McpSource) {
        Copy-Item -Path $McpSource -Destination "$SkillsStagingDir/excel-mcp" -Recurse
        Copy-SharedReferences -SkillPath "$SkillsStagingDir/excel-mcp" -SkillName "excel-mcp"
    } else {
        Write-Warning "excel-mcp skill not found"
    }

    # Copy excel-cli skill
    $CliSource = Join-Path $SkillsDir "excel-cli"
    if (Test-Path $CliSource) {
        Copy-Item -Path $CliSource -Destination "$SkillsStagingDir/excel-cli" -Recurse
        Copy-SharedReferences -SkillPath "$SkillsStagingDir/excel-cli" -SkillName "excel-cli"
        # Generate CLI command reference from excelcli --help
        Generate-CliReference -SkillPath "$SkillsStagingDir/excel-cli"
    } else {
        Write-Warning "excel-cli skill not found"
    }

    # Copy skills README to root of package
    $SkillsReadme = Join-Path $SkillsDir "README.md"
    if (Test-Path $SkillsReadme) {
        Copy-Item -Path $SkillsReadme -Destination $StagingDir
    }

    # Create ZIP archive
    $ZipName = "excel-skills-v$Version.zip"
    $ZipPath = Join-Path $OutputPath $ZipName

    if (Test-Path $ZipPath) {
        Remove-Item $ZipPath -Force
    }

    Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal
    Write-Host "  Created: $ZipName" -ForegroundColor Green

} finally {
    if (Test-Path $StagingDir) {
        Remove-Item $StagingDir -Recurse -Force
    }
}

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
    name = "excel-skills"
    version = $Version
    description = "Excel MCP Server Agent Skills for AI coding assistants"
    platforms = @("github-copilot", "claude-code", "cursor", "windsurf", "gemini-cli", "goose", "codex", "opencode", "amp", "kilo", "roo", "trae")
    skills = @(
        @{
            name = "excel-mcp"
            path = "skills/excel-mcp"
            description = "MCP Server skill - for conversational AI (Claude Desktop, VS Code Chat)"
            target = "MCP Server"
        }
        @{
            name = "excel-cli"
            path = "skills/excel-cli"
            description = "CLI skill - for coding agents (Copilot, Cursor, Windsurf)"
            target = "CLI Tool"
        }
    )
    installation = @{
        npx = "npx skills add sbroenne/mcp-server-excel"
        selectSkill = "npx skills add sbroenne/mcp-server-excel --skill excel-cli"
        installBoth = "npx skills add sbroenne/mcp-server-excel --skill '*'"
    }
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

Write-Host ""
Write-Host "Installation:" -ForegroundColor Cyan
Write-Host "  npx skills add sbroenne/mcp-server-excel" -ForegroundColor White
Write-Host "  (users will be prompted to select excel-cli, excel-mcp, or both)"

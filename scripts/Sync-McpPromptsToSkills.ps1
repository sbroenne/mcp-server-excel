<#
.SYNOPSIS
    Synchronizes MCP Server prompts to skills/shared directory.

.DESCRIPTION
    Copies tool-specific guidance files from MCP Server prompts (source of truth)
    to skills/shared directory for agent skills distribution.
    
    MCP Server prompts are embedded resources and served via MCP protocol.
    Agent skills need standalone markdown files for offline use.

.PARAMETER Verify
    Verify synchronization without copying files.

.EXAMPLE
    ./Sync-McpPromptsToSkills.ps1
    
.EXAMPLE
    ./Sync-McpPromptsToSkills.ps1 -Verify
#>
param(
    [switch]$Verify
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$McpPromptsDir = Join-Path $RepoRoot "src/ExcelMcp.McpServer/Prompts/Content"
$SkillsSharedDir = Join-Path $RepoRoot "skills/shared"

# Files to sync (tool-specific guidance only, not agent skill-specific files)
$FilesToSync = @(
    "excel_chart.md",
    "excel_conditionalformat.md",
    "excel_datamodel.md",
    "excel_powerquery.md",
    "excel_range.md",
    "excel_slicer.md",
    "excel_table.md",
    "excel_worksheet.md"
)

function Get-FileHashSafe {
    param([string]$Path)
    if (Test-Path $Path) {
        return (Get-FileHash -Path $Path -Algorithm SHA256).Hash
    }
    return $null
}

if ($Verify) {
    Write-Host "Verifying synchronization..." -ForegroundColor Cyan
    Write-Host ""
    
    $allSynced = $true
    foreach ($file in $FilesToSync) {
        $mcpFile = Join-Path $McpPromptsDir $file
        $sharedFile = Join-Path $SkillsSharedDir $file
        
        $mcpHash = Get-FileHashSafe $mcpFile
        $sharedHash = Get-FileHashSafe $sharedFile
        
        if ($mcpHash -and $sharedHash) {
            if ($mcpHash -eq $sharedHash) {
                Write-Host "✓ $file - SYNCHRONIZED" -ForegroundColor Green
            } else {
                Write-Host "✗ $file - OUT OF SYNC" -ForegroundColor Red
                $allSynced = $false
            }
        } elseif (-not $mcpHash) {
            Write-Host "⚠ $file - MISSING IN MCP PROMPTS" -ForegroundColor Yellow
            $allSynced = $false
        } elseif (-not $sharedHash) {
            Write-Host "⚠ $file - MISSING IN SKILLS/SHARED" -ForegroundColor Yellow
            $allSynced = $false
        }
    }
    
    Write-Host ""
    if ($allSynced) {
        Write-Host "All files are synchronized!" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "Some files are out of sync. Run without -Verify to synchronize." -ForegroundColor Yellow
        exit 1
    }
}

# Sync mode (copy files)
Write-Host "Synchronizing MCP Server prompts to skills/shared..." -ForegroundColor Cyan
Write-Host "Source: $McpPromptsDir" -ForegroundColor Gray
Write-Host "Target: $SkillsSharedDir" -ForegroundColor Gray
Write-Host ""

$copiedCount = 0
$skippedCount = 0

foreach ($file in $FilesToSync) {
    $mcpFile = Join-Path $McpPromptsDir $file
    $sharedFile = Join-Path $SkillsSharedDir $file
    
    if (-not (Test-Path $mcpFile)) {
        Write-Host "⚠ Skipping $file - not found in MCP prompts" -ForegroundColor Yellow
        $skippedCount++
        continue
    }
    
    # Check if already synchronized
    $mcpHash = Get-FileHashSafe $mcpFile
    $sharedHash = Get-FileHashSafe $sharedFile
    
    if ($mcpHash -eq $sharedHash) {
        Write-Host "✓ $file - already synchronized" -ForegroundColor Gray
        continue
    }
    
    # Copy file
    Copy-Item -Path $mcpFile -Destination $sharedFile -Force
    Write-Host "→ Copied $file" -ForegroundColor Green
    $copiedCount++
}

Write-Host ""
Write-Host "Synchronization complete!" -ForegroundColor Green
Write-Host "  Copied: $copiedCount file(s)" -ForegroundColor Green
if ($skippedCount -gt 0) {
    Write-Host "  Skipped: $skippedCount file(s)" -ForegroundColor Yellow
}

#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validates plugin README overlay quality before commit

.DESCRIPTION
    Checks plugin README overlays in .github/plugins/*/README.md for:
    - Minimum content length (detects thin/stub content)
    - Required sections for published marketplace plugins
    - Prevents shipping incomplete plugin documentation
    
    This gate catches thin README overlays that would overwrite richer
    published templates, preventing users from seeing incomplete docs.

.EXAMPLE
    .\check-plugin-readmes.ps1
    
.NOTES
    Part of pre-commit validation. Added after v1.8.44 to prevent
    shipping stub plugin READMEs to marketplace.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$pluginOverlayDir = Join-Path $rootDir ".github\plugins"

Write-Host "Validating plugin README overlays..." -ForegroundColor Cyan
Write-Host ""

# Minimum content requirements for published plugin READMEs
$MINIMUM_LINES = 40  # Stub content is usually < 20 lines; full content is 80+
$REQUIRED_SECTIONS = @(
    "# Excel",           # Plugin title
    "## Prerequisites",  # Requirements
    "## Installation"    # How to install
    # Optional: "## What's Included", "## Notes", "## Support"
)

$violations = @()

# Find all plugin README overlays (skip marketplace-repo README - that's repo-level docs)
$pluginReadmes = Get-ChildItem -Path $pluginOverlayDir -Filter "README.md" -Recurse | 
    Where-Object { $_.DirectoryName -notmatch "marketplace-repo" }

if ($pluginReadmes.Count -eq 0) {
    Write-Host "No plugin README overlays found - validation skipped" -ForegroundColor Yellow
    exit 0
}

foreach ($readme in $pluginReadmes) {
    $pluginName = Split-Path -Parent $readme.DirectoryName | Split-Path -Leaf
    $pluginPath = $readme.DirectoryName | Split-Path -Leaf
    $relativePath = $readme.FullName.Replace("$rootDir\", "")
    
    Write-Host "Checking: $relativePath" -ForegroundColor White
    
    $content = Get-Content $readme.FullName -Raw
    $lines = Get-Content $readme.FullName
    
    # Check 1: Minimum line count
    $lineCount = $lines.Count
    if ($lineCount -lt $MINIMUM_LINES) {
        $violations += [PSCustomObject]@{
            File = $relativePath
            Issue = "Insufficient content"
            Details = "Only $lineCount lines (minimum: $MINIMUM_LINES). This appears to be stub content."
            Severity = "ERROR"
        }
        Write-Host "  ❌ Too short: $lineCount lines (minimum: $MINIMUM_LINES)" -ForegroundColor Red
    } else {
        Write-Host "  ✓ Length OK: $lineCount lines" -ForegroundColor Green
    }
    
    # Check 2: Required sections
    $missingSections = @()
    foreach ($section in $REQUIRED_SECTIONS) {
        if ($content -notmatch [regex]::Escape($section)) {
            $missingSections += $section
        }
    }
    
    if ($missingSections.Count -gt 0) {
        $violations += [PSCustomObject]@{
            File = $relativePath
            Issue = "Missing required sections"
            Details = "Missing: $($missingSections -join ', ')"
            Severity = "ERROR"
        }
        Write-Host "  ❌ Missing sections: $($missingSections -join ', ')" -ForegroundColor Red
    } else {
        Write-Host "  ✓ All required sections present" -ForegroundColor Green
    }
    
    Write-Host ""
}

# Report results
if ($violations.Count -eq 0) {
    Write-Host "All plugin README overlays validated!" -ForegroundColor Green
    Write-Host "  Files checked: $($pluginReadmes.Count)" -ForegroundColor Cyan
    exit 0
} else {
    Write-Host "VALIDATION FAILED: Plugin README overlays have issues" -ForegroundColor Red
    Write-Host ""
    
    foreach ($violation in $violations) {
        Write-Host "[$($violation.Severity)] $($violation.File)" -ForegroundColor Red
        Write-Host "  Issue: $($violation.Issue)" -ForegroundColor Yellow
        Write-Host "  $($violation.Details)" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "WHY THIS MATTERS:" -ForegroundColor Yellow
    Write-Host "  Plugin README overlays overwrite published templates during build." -ForegroundColor White
    Write-Host "  Stub/thin content means users see incomplete docs in the marketplace." -ForegroundColor White
    Write-Host ""
    Write-Host "TO FIX:" -ForegroundColor Yellow
    Write-Host "  1. Enrich the README overlay with full content (80+ lines)" -ForegroundColor White
    Write-Host "  2. Include all required sections (Prerequisites, Installation, etc.)" -ForegroundColor White
    Write-Host "  3. Or remove the overlay to use the published template as-is" -ForegroundColor White
    Write-Host ""
    
    exit 1
}

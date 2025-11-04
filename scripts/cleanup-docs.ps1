# Repository Documentation Cleanup Script
# Run this after extracting information from temporary files

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Repository Documentation Cleanup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Check for temporary files
Write-Host "Step 1: Checking for temporary files to delete..." -ForegroundColor Yellow
Write-Host ""

$tempFiles = @(
    "CODEQL-FIXES-SUMMARY.md",
    "CODEQL-SUPPRESSION-VERIFICATION.md",
    "docs\PR_SUMMARY.md",
    "vscode-extension\SUMMARY.md",
    "tests\ExcelMcp.Core.Tests\docs\VBA-TEST-EXCLUSION-SUMMARY.md"
)

$foundFiles = @()
foreach ($file in $tempFiles) {
    if (Test-Path $file) {
        $foundFiles += $file
        Write-Host "  ✓ Found: $file" -ForegroundColor Green
    } else {
        Write-Host "  ✗ Not found: $file (already deleted?)" -ForegroundColor Gray
    }
}

Write-Host ""

# Step 2: Prompt for confirmation
if ($foundFiles.Count -eq 0) {
    Write-Host "No temporary files found. Cleanup already complete!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Checking for any other temporary files..." -ForegroundColor Yellow
    
    $otherTemps = Get-ChildItem -Recurse -Filter "*.md" | 
        Where-Object { 
            $_.FullName -notlike "*\node_modules\*" -and 
            $_.FullName -notlike "*\.git\*" -and 
            ($_.Name -match "(SUMMARY|FIX|TESTS|DOCS|MIGRATION|AUDIT|PLAN)") -and
            $_.Name -ne "CLEANUP-PROPOSAL.md" -and
            $_.Name -ne "CLEANUP-IMPLEMENTATION-SUMMARY.md" -and
            $_.Name -ne "PR_SUMMARY.md" -and
            $_.Name -ne "TEST-NAMING-STANDARD.md" -and
            $_.Name -ne "DATA-MODEL-SETUP.md" -and
            $_.Name -ne "AZURE_SELFHOSTED_RUNNER_SETUP.md" -and
            $_.Name -ne "PRE-COMMIT-SETUP.md"
        }
    
    if ($otherTemps.Count -eq 0) {
        Write-Host "✓ No other temporary files found!" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "Found other potential temporary files:" -ForegroundColor Yellow
        $otherTemps | ForEach-Object { Write-Host "  - $($_.FullName)" -ForegroundColor Cyan }
    }
    
    exit 0
}

Write-Host "Found $($foundFiles.Count) temporary files to delete." -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠️  WARNING: Before deleting, ensure you have:" -ForegroundColor Red
Write-Host "   1. Extracted valuable information from each file" -ForegroundColor Red
Write-Host "   2. Merged information into permanent documentation" -ForegroundColor Red
Write-Host "   3. Verified permanent docs contain all necessary details" -ForegroundColor Red
Write-Host ""

$confirmation = Read-Host "Have you completed all extractions? (yes/no)"

if ($confirmation -ne "yes") {
    Write-Host ""
    Write-Host "Cleanup cancelled. Please extract information first." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Extraction targets:" -ForegroundColor Cyan
    Write-Host "  1. CODEQL-FIXES-SUMMARY.md → Git history (no extraction needed)" -ForegroundColor Gray
    Write-Host "  2. CODEQL-SUPPRESSION-VERIFICATION.md → No extraction needed" -ForegroundColor Gray
    Write-Host "  3. docs\PR_SUMMARY.md → docs\AZURE_SELFHOSTED_RUNNER_SETUP.md" -ForegroundColor Gray
    Write-Host "  4. vscode-extension\SUMMARY.md → vscode-extension\README.md" -ForegroundColor Gray
    Write-Host "  5. tests\...\VBA-TEST-EXCLUSION-SUMMARY.md → tests\README.md" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

# Step 3: Delete files
Write-Host ""
Write-Host "Step 2: Deleting temporary files..." -ForegroundColor Yellow
Write-Host ""

$deletedCount = 0
foreach ($file in $foundFiles) {
    try {
        Remove-Item $file -Force -ErrorAction Stop
        Write-Host "  ✓ Deleted: $file" -ForegroundColor Green
        $deletedCount++
    } catch {
        Write-Host "  ✗ Failed to delete: $file" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Deleted $deletedCount of $($foundFiles.Count) files." -ForegroundColor Cyan

# Step 4: Delete proposal files
Write-Host ""
Write-Host "Step 3: Delete proposal files? (CLEANUP-PROPOSAL.md, CLEANUP-IMPLEMENTATION-SUMMARY.md)" -ForegroundColor Yellow
$deleteProposal = Read-Host "Delete proposal files now? (yes/no)"

if ($deleteProposal -eq "yes") {
    $proposalFiles = @("CLEANUP-PROPOSAL.md", "CLEANUP-IMPLEMENTATION-SUMMARY.md")
    foreach ($file in $proposalFiles) {
        if (Test-Path $file) {
            Remove-Item $file -Force
            Write-Host "  ✓ Deleted: $file" -ForegroundColor Green
        }
    }
}

# Step 5: Verify cleanup
Write-Host ""
Write-Host "Step 4: Verifying cleanup..." -ForegroundColor Yellow
Write-Host ""

$remainingTemps = Get-ChildItem -Recurse -Filter "*.md" | 
    Where-Object { 
        $_.FullName -notlike "*\node_modules\*" -and 
        $_.FullName -notlike "*\.git\*" -and 
        ($_.Name -match "(SUMMARY|FIX|TESTS|DOCS|MIGRATION|AUDIT|PLAN)") -and
        $_.Name -ne "TEST-NAMING-STANDARD.md" -and
        $_.Name -ne "DATA-MODEL-SETUP.md" -and
        $_.Name -ne "AZURE_SELFHOSTED_RUNNER_SETUP.md" -and
        $_.Name -ne "PRE-COMMIT-SETUP.md" -and
        $_.Name -ne "CLEANUP-PROPOSAL.md" -and
        $_.Name -ne "CLEANUP-IMPLEMENTATION-SUMMARY.md"
    }

if ($remainingTemps.Count -eq 0) {
    Write-Host "✓ Cleanup complete! No temporary files remaining." -ForegroundColor Green
} else {
    Write-Host "⚠️  Warning: Found $($remainingTemps.Count) potential temporary files:" -ForegroundColor Yellow
    $remainingTemps | ForEach-Object { 
        Write-Host "  - $($_.FullName)" -ForegroundColor Cyan 
    }
    Write-Host ""
    Write-Host "Review these files manually to determine if they should be kept or deleted." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "1. Review changes in permanent documentation" -ForegroundColor White
Write-Host "2. Run: git add -A" -ForegroundColor White
Write-Host "3. Run: git commit -m 'docs: cleanup temporary files'" -ForegroundColor White
Write-Host "4. Run: git push" -ForegroundColor White
Write-Host ""

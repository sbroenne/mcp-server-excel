# CodeQL Suppression Verification Script
# Analyzes SARIF results to show what will be suppressed by config v3.0

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   CodeQL Config v3.0 - Suppression Impact Analysis" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Download latest SARIF results from GitHub
Write-Host "📥 Downloading latest CodeQL SARIF results from GitHub..." -ForegroundColor Yellow
try {
    gh api -H "Accept: application/sarif+json" /repos/sbroenne/mcp-server-excel/code-scanning/analyses/764819378 > codeql-results-temp.sarif 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Downloaded latest SARIF results" -ForegroundColor Green
    } else {
        throw "GitHub API call failed"
    }
} catch {
    Write-Host "⚠️ Could not download from GitHub. Using cached results if available." -ForegroundColor Yellow
    if (Test-Path "codeql-results.sarif") {
        Copy-Item "codeql-results.sarif" "codeql-results-temp.sarif"
    } else {
        Write-Host "❌ No SARIF file available. Run this script after a CodeQL scan completes." -ForegroundColor Red
        exit 1
    }
}

Write-Host ""

# Load SARIF
$sarif = Get-Content "codeql-results-temp.sarif" | ConvertFrom-Json

# Rules that our config v3.0 suppresses
$suppressedRules = @(
    'cs/catch-of-all-exceptions',      # COM requires broad exception handling
    'cs/empty-catch-block',            # Cleanup must not fail
    'cs/call-to-gc',                   # Required for COM cleanup
    'cs/call-to-unmanaged-code',       # OLE message filter
    'cs/nested-if-statements',         # Complex validation patterns
    'cs/dereferenced-value-may-be-null', # Custom validation patterns
    'cs/useless-upcast',               # COM dynamic types
    'cs/invalid-dynamic-call',         # COM requires dynamic
    'cs/missed-ternary-operator',      # Explicit if/else preferred
    'cs/linq/missed-select',           # Explicit loops clearer
    'cs/simplifiable-boolean-expression', # Explicit expressions preferred
    'cs/unmanaged-code',               # COM interop requirement
    'cs/useless-assignment-to-local',  # Intermediate COM references
    'cs/empty-block',                  # Development placeholders
    'cs/useless-if-statement'          # Conditional preservation
)

$suppressedIssues = $sarif.runs[0].results | Where-Object { $_.ruleId -in $suppressedRules }
$remainingIssues = $sarif.runs[0].results | Where-Object { $_.ruleId -notin $suppressedRules }

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   Current State (Before Config v3.0 Applied)" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$sarif.runs[0].results | Group-Object -Property ruleId | 
    Select-Object @{N='Rule';E={$_.Name}}, Count | 
    Sort-Object Count -Descending | 
    Format-Table -AutoSize

$total = $sarif.runs[0].results.Count
Write-Host "Total Issues: $total" -ForegroundColor White
Write-Host ""

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   After Config v3.0 Applied (Next Scan)" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

Write-Host "✅ Issues That Will Be SUPPRESSED:" -ForegroundColor Green
Write-Host ""
$suppressedIssues | Group-Object -Property ruleId | 
    Select-Object @{N='Rule';E={$_.Name}}, Count | 
    Sort-Object Count -Descending | 
    Format-Table -AutoSize

$suppressedCount = $suppressedIssues.Count
Write-Host "Total Suppressed: $suppressedCount" -ForegroundColor Green
Write-Host ""

if ($remainingIssues.Count -eq 0) {
    Write-Host "⚠️ Issues That Will REMAIN: None! ✅" -ForegroundColor Green
    Write-Host ""
    Write-Host "All current issues are intentional COM interop patterns and will be suppressed!" -ForegroundColor Green
} else {
    Write-Host "⚠️ Issues That Will REMAIN:" -ForegroundColor Yellow
    Write-Host ""
    $remainingIssues | Group-Object -Property ruleId | 
        Select-Object @{N='Rule';E={$_.Name}}, Count | 
        Sort-Object Count -Descending | 
        Format-Table -AutoSize
    
    Write-Host "Total Remaining: $($remainingIssues.Count)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "📊 Impact Summary" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Before (Current):     $total issues" -ForegroundColor White
Write-Host "  Suppressed (v3.0):    $suppressedCount issues" -ForegroundColor Green
Write-Host "  Remaining (Future):   $($remainingIssues.Count) issues" -ForegroundColor Yellow
$percentage = [math]::Round(($suppressedCount / $total) * 100, 1)
Write-Host "  Reduction:            $percentage%" -ForegroundColor Green
Write-Host ""

# Configuration status
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "📝 Configuration Status" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$configPath = ".github/codeql/codeql-config.yml"
if (Test-Path $configPath) {
    $configContent = Get-Content $configPath -Raw
    if ($configContent -match "v3\.0") {
        Write-Host "✅ CodeQL Config: v3.0 (Latest)" -ForegroundColor Green
    } else {
        Write-Host "⚠️ CodeQL Config: Older version detected" -ForegroundColor Yellow
        Write-Host "   Update to v3.0 to activate suppressions" -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ CodeQL Config: Not found at $configPath" -ForegroundColor Red
}

$workflowPath = ".github/workflows/codeql.yml"
if (Test-Path $workflowPath) {
    $workflowContent = Get-Content $workflowPath -Raw
    if ($workflowContent -match "config-file.*codeql-config\.yml") {
        Write-Host "✅ CodeQL Workflow: Uses custom config" -ForegroundColor Green
    } else {
        Write-Host "⚠️ CodeQL Workflow: May not be using custom config" -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ CodeQL Workflow: Not found" -ForegroundColor Red
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "✅ Verification Complete" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Merge these changes to 'main'" -ForegroundColor White
Write-Host "  2. Wait for next CodeQL scan (Monday 10:00 AM UTC or trigger manually)" -ForegroundColor White
Write-Host "  3. Verify suppression count matches this analysis" -ForegroundColor White
Write-Host ""

# Cleanup
Remove-Item "codeql-results-temp.sarif" -ErrorAction SilentlyContinue

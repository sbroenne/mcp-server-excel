#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook to check for COM object leaks, Core Commands coverage, naming consistency, Success flag violations, and MCP Server functionality

.DESCRIPTION
    Runs five checks before allowing commits:
    1. COM leak checker - ensures no Excel COM objects are leaked
    2. Coverage audit - ensures 100% Core Commands are exposed via MCP Server
    3. Naming consistency - ensures enum names match Core method names exactly
    4. Success flag validation - ensures Success=true never paired with ErrorMessage (Rule 0)
    5. Smoke test - validates all 11 MCP tools work correctly

    Ensures code quality and prevents regression.

.EXAMPLE
    .\pre-commit.ps1

.NOTES
    This script is called by the Git pre-commit hook.
    To install: Copy .git/hooks/pre-commit (bash) or configure Git to use this PowerShell version.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

# CRITICAL: Check branch FIRST - never commit directly to main (Rule 6)
Write-Host "🔍 Checking current branch..." -ForegroundColor Cyan
$currentBranch = git branch --show-current

if ($currentBranch -eq "main") {
    Write-Host ""
    Write-Host "❌ BLOCKED: Cannot commit directly to 'main' branch!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Rule 6: All Changes Via Pull Requests" -ForegroundColor Yellow
    Write-Host "   'Never commit to main. Create feature branch → PR → CI/CD + review → merge.'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   To fix:" -ForegroundColor Cyan
    Write-Host "   1. git stash                                    # Save your changes" -ForegroundColor White
    Write-Host "   2. git checkout -b feature/your-feature-name    # Create feature branch" -ForegroundColor White
    Write-Host "   3. git stash pop                                # Restore changes" -ForegroundColor White
    Write-Host "   4. git add <files>                              # Stage changes" -ForegroundColor White
    Write-Host "   5. git commit -m 'your message'                 # Commit to feature branch" -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "✅ Branch check passed - on '$currentBranch' (not main)" -ForegroundColor Green
Write-Host ""

Write-Host "🔍 Checking for COM object leaks..." -ForegroundColor Cyan

try {
    $leakCheckScript = Join-Path $rootDir "scripts\check-com-leaks.ps1"
    & $leakCheckScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ COM object leaks detected! Fix them before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ COM leak check passed" -ForegroundColor Green
}
catch {
    Write-Host "⚠️  Error running COM leak check: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   Continuing with coverage audit..." -ForegroundColor Gray
}

Write-Host ""
Write-Host "🔍 Checking Core Commands coverage and naming..." -ForegroundColor Cyan

try {
    $auditScript = Join-Path $rootDir "scripts\audit-core-coverage.ps1"
    & $auditScript -CheckNaming -FailOnGaps

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ Coverage or naming issues detected!" -ForegroundColor Red
        Write-Host "   All Core methods must be exposed via MCP Server with matching names." -ForegroundColor Red
        Write-Host "   Fix the issues before committing (add/rename enum values and mappings)." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ Coverage and naming checks passed - 100% coverage with consistent names" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running coverage audit: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔍 Checking MCP actions have Core implementations..." -ForegroundColor Cyan

try {
    $mcpCoreScript = Join-Path $rootDir "scripts\check-mcp-core-implementations.ps1"
    & $mcpCoreScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ MCP actions without Core implementations detected!" -ForegroundColor Red
        Write-Host "   All enum actions must have matching Core Command methods." -ForegroundColor Red
        Write-Host "   Fix the issues before committing (remove enum or implement method)." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ MCP-Core implementation check passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running MCP-Core implementation check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔍 Checking Success flag violations (Rule 0)..." -ForegroundColor Cyan

try {
    $successFlagScript = Join-Path $rootDir "scripts\check-success-flag.ps1"
    & $successFlagScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ Success flag violations detected!" -ForegroundColor Red
        Write-Host "   CRITICAL: Success=true with ErrorMessage confuses LLMs and causes data corruption." -ForegroundColor Red
        Write-Host "   Fix the violations before committing (add Success=false in catch blocks)." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ Success flag check passed - all flags match reality" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running success flag check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔍 Running CLI actions audit..." -ForegroundColor Cyan

try {
    $cliAuditScript = Join-Path $rootDir "scripts\audit-cli-actions.ps1"
    & $cliAuditScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "❌ CLI actions audit failed!" -ForegroundColor Red
        Write-Host "   CLI action catalog must match Core action enums." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "✅ CLI actions audit passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running CLI actions audit: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔍 Running CLI workflow smoke test..." -ForegroundColor Cyan

try {
    $cliWorkflowScript = Join-Path $rootDir "scripts\Test-CliWorkflow.ps1"
    $cliWorkflowOutput = & $cliWorkflowScript 2>&1 | Out-String
    $cliWorkflowExitCode = $LASTEXITCODE

    if ($cliWorkflowExitCode -ne 0) {
        Write-Host ""
        Write-Host "❌ CLI workflow smoke test failed!" -ForegroundColor Red
        Write-Host "   This test validates the end-to-end CLI workflow." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $cliWorkflowOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "✅ CLI workflow smoke test passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running CLI workflow smoke test: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔍 Running MCP Server smoke test..." -ForegroundColor Cyan

try {
    # Run the smoke test - validates all MCP tools work correctly
    $smokeTestFilter = "FullyQualifiedName~McpServerSmokeTests.SmokeTest_AllTools_E2EWorkflow"

    Write-Host "   dotnet test --filter `"$smokeTestFilter`"" -ForegroundColor Gray
    
    # Capture output to verify tests actually ran (dotnet test returns 0 even if no tests match!)
    $testOutput = dotnet test --filter $smokeTestFilter --verbosity minimal 2>&1 | Out-String
    $testExitCode = $LASTEXITCODE
    
    # Check if any tests actually passed (critical - filter typos cause silent failures!)
    # Note: "No test matches" appears for projects without the test, so we check for "Passed"
    if (-not ($testOutput -match "Passed!.*Passed:\s*[1-9]")) {
        Write-Host ""
        Write-Host "❌ CRITICAL: No smoke tests passed! Filter may have matched zero tests." -ForegroundColor Red
        Write-Host "   Filter: $smokeTestFilter" -ForegroundColor Yellow
        Write-Host "   This likely means the test was renamed or deleted." -ForegroundColor Yellow
        Write-Host "   Verify the test exists: McpServerSmokeTests.SmokeTest_AllTools_E2EWorkflow" -ForegroundColor Yellow
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }
    
    if ($testExitCode -ne 0) {
        Write-Host ""
        Write-Host "❌ MCP Server smoke test failed! Core functionality is broken." -ForegroundColor Red
        Write-Host "   This test validates all MCP tools work correctly." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "✅ MCP Server smoke test passed - all tools functional" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "❌ Error running smoke test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Ensure Excel is installed and accessible." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "✅ All pre-commit checks passed!" -ForegroundColor Green
exit 0

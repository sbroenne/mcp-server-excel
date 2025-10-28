#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Automates conversion of integration tests from synchronous API to async batch-of-one pattern.

.DESCRIPTION
    This script converts ExcelMcp integration tests to use the new async batch API:
    1. Adds using statements for ExcelSession
    2. Converts method signatures from void to async Task
    3. Wraps command calls in batch-of-one pattern
    4. Handles special cases (file creation, set-then-get, non-existent files)

.PARAMETER TestFile
    Path to the test file to convert. If not specified, processes all integration test files.

.PARAMETER WhatIf
    Shows what would be changed without making modifications.

.EXAMPLE
    .\Convert-TestsToBatchPattern.ps1 -TestFile "tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQueryCommandsTests.cs"

.EXAMPLE
    .\Convert-TestsToBatchPattern.ps1 -WhatIf
#>

param(
    [string]$TestFile,
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"

# Test file patterns to process
$testFiles = if ($TestFile) {
    @($TestFile)
} else {
    @(
        "tests/ExcelMcp.Core.Tests/Integration/Commands/CoreConnectionCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/Commands/CoreConnectionCommandsExtendedTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/Commands/SheetCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/Commands/SetupCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/Commands/VbaTrustDetectionTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/PowerQuery/PowerQueryWorkflowGuidanceTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/PowerQuery/PowerQueryPrivacyLevelTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/PowerQuery/PowerQueryCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/DataModel/DataModelCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/DataModel/DataModelTomCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/Commands/FileCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/Integration/Commands/ParameterCommandsTests.cs",
        "tests/ExcelMcp.Core.Tests/RoundTrip/IntegrationWorkflowTests.cs",
        "tests/ExcelMcp.Core.Tests/RoundTrip/ConnectionWorkflowTests.cs",
        "tests/ExcelMcp.Core.Tests/RoundTrip/ScriptCommandsRoundTripTests.cs"
    )
}

function Add-UsingStatement {
    param([string]$content)
    
    if ($content -notmatch "using Sbroenne\.ExcelMcp\.Core\.Session;") {
        # Find the last using statement
        $lines = $content -split "`n"
        $lastUsingIndex = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^using ") {
                $lastUsingIndex = $i
            }
        }
        
        if ($lastUsingIndex -ge 0) {
            $lines = $lines[0..$lastUsingIndex] + "using Sbroenne.ExcelMcp.Core.Session;" + $lines[($lastUsingIndex + 1)..($lines.Count - 1)]
            $content = $lines -join "`n"
        }
    }
    
    return $content
}

function Convert-MethodSignature {
    param([string]$content)
    
    # Convert public void TestMethod() to public async Task TestMethod()
    $content = $content -replace '(\[Fact\][\r\n\s]+public) void (\w+\(\))', '$1 async Task $2'
    
    return $content
}

function Convert-CommandCalls {
    param([string]$content, [string]$fileName)
    
    # This is complex - we'll do basic transformations
    
    # Pattern 1: Read operations without save
    # _commands.Method(filePath, ...) -> await _commands.MethodAsync(batch, ...)
    
    # Pattern 2: Write operations with save
    # var result = _commands.Method(filePath, ...);
    # ->
    # await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    # var result = await _commands.MethodAsync(batch, ...);
    # await batch.SaveAsync();
    
    # This requires more sophisticated parsing - let's mark what needs manual attention
    
    return $content
}

function Convert-TestFile {
    param([string]$filePath)
    
    Write-Host "Processing: $filePath" -ForegroundColor Cyan
    
    if (-not (Test-Path $filePath)) {
        Write-Warning "File not found: $filePath"
        return
    }
    
    $content = Get-Content $filePath -Raw
    $originalContent = $content
    
    # Step 1: Add using statement
    $content = Add-UsingStatement $content
    
    # Step 2: Convert method signatures
    $content = Convert-MethodSignature $content
    
    # Step 3: Command call conversion (basic patterns)
    $content = Convert-CommandCalls $content $filePath
    
    if ($content -ne $originalContent) {
        if ($WhatIf) {
            Write-Host "  Would update: $filePath" -ForegroundColor Yellow
        } else {
            Set-Content $filePath $content -NoNewline
            Write-Host "  Updated: $filePath" -ForegroundColor Green
        }
    } else {
        Write-Host "  No changes needed" -ForegroundColor Gray
    }
}

# Main execution
Write-Host "Test File Batch Conversion Script" -ForegroundColor Magenta
Write-Host "=================================" -ForegroundColor Magenta
Write-Host ""

if ($WhatIf) {
    Write-Host "Running in WhatIf mode - no files will be modified" -ForegroundColor Yellow
    Write-Host ""
}

$totalFiles = 0
foreach ($file in $testFiles) {
    Convert-TestFile $file
    $totalFiles++
}

Write-Host ""
Write-Host "Processed $totalFiles file(s)" -ForegroundColor Magenta

Write-Host ""
Write-Host "IMPORTANT: This script handles basic transformations only." -ForegroundColor Yellow
Write-Host "Manual review and completion required for:" -ForegroundColor Yellow
Write-Host "  1. Wrapping method calls in batch-of-one pattern" -ForegroundColor Yellow
Write-Host "  2. Adding await batch.SaveAsync() after write operations" -ForegroundColor Yellow
Write-Host "  3. Separating batch scopes for set-then-get operations" -ForegroundColor Yellow
Write-Host "  4. Converting non-existent file tests to Assert.ThrowsAsync" -ForegroundColor Yellow
Write-Host ""
Write-Host "After running this script, build the test project to find remaining errors:" -ForegroundColor Cyan
Write-Host "  dotnet build tests/ExcelMcp.Core.Tests" -ForegroundColor Cyan

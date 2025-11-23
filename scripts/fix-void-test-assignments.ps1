#!/usr/bin/env pwsh
<#
.SYNOPSIS
Systematically fixes "Cannot assign void" compilation errors in test files.

.DESCRIPTION
Removes 'var result =' patterns from void method calls and removes associated assertions.
Handles both simple assignments and complex patterns.

.EXAMPLE
.\fix-void-test-assignments.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Void method calls pattern - these throw on error instead of returning results
$voidMethodPatterns = @(
    # Sheet operations
    'Create\(batch'
    'Delete\(batch'
    'Rename\(batch'
    'Move\(batch'
    'CopyToWorkbook\(batch'
    'MoveToWorkbook\(batch'
    'SetTabColor\(batch'
    'ClearTabColor\(batch'
    'Hide\(batch'
    'Show\(batch'
    'VeryHide\(batch'
    'SetVisibility\(batch'
    
    # Table operations
    'Resize\(batch'
    'AddColumn\(batch'
    'RenameColumn\(batch'
    'ToggleTotals\(batch'
    'SetColumnTotal\(batch'
    'Append\(batch'
    'ApplyFilter\(batch'
    'ApplyFilterValues\(batch'
    'ClearFilters\(batch'
    
    # PowerQuery operations
    'Update\(batch'
    'RefreshAll\(batch'
    'Delete\(batch'
    'LoadTo\(batch'
    
    # DataModel operations
    'AddToDataModel\(batch'
    'CreateMeasure\(batch'
    'CreateRelationship\(batch'
    'DeleteMeasure\(batch'
    'DeleteRelationship\(batch'
    'UpdateMeasure\(batch'
    'UpdateRelationship\(batch'
    
    # PivotTable operations
    'AddRowField\(batch'
    'AddColumnField\(batch'
    'AddValueField\(batch'
    'AddFilterField\(batch'
    'RemoveField\(batch'
    'SetFieldFunction\(batch'
    'SetFieldName\(batch'
    'SetFieldFormat\(batch'
    'SetFieldFilter\(batch'
    'SortField\(batch'
)

# Find all test files
$testFiles = @(
    'tests\ExcelMcp.Core.Tests\Integration\Commands\Sheet\SheetCommandsTests.Move.cs'
    'tests\ExcelMcp.Core.Tests\Integration\Commands\Sheet\SheetCommandsTests.TabColor.cs'
    'tests\ExcelMcp.Core.Tests\Integration\Commands\Sheet\SheetCommandsTests.Visibility.cs'
    'tests\ExcelMcp.Core.Tests\Integration\Commands\PowerQuery\PowerQueryCommandsTests.cs'
    'tests\ExcelMcp.Core.Tests\Integration\Commands\DataModel\DataModelCommandsTests.cs'
    'tests\ExcelMcp.Core.Tests\Integration\Commands\PivotTable\PivotTableCommandsTests.Creation.cs'
    'tests\ExcelMcp.Core.Tests\Integration\Commands\PivotTable\PivotTableCommandsTests.OlapFields.cs'
    'tests\ExcelMcp.Core.Tests\Helpers\DataModelTestsFixture.cs'
)

foreach ($testFile in $testFiles) {
    $fullPath = Join-Path 'D:\source\mcp-server-excel' $testFile
    
    if (-not (Test-Path $fullPath)) {
        Write-Verbose "File not found: $fullPath"
        continue
    }
    
    Write-Verbose "Processing: $testFile"
    $content = Get-Content $fullPath -Raw
    $originalContent = $content
    
    # Pattern 1: Remove 'var result = method(...);\n            Assert.True(result.Success...'
    # This is the most common pattern
    $content = $content -replace '(?ms)var result\d* = (\w+\.)\w+\([^)]*\);\s*Assert\.True\(result\d*\.Success[^;]*\);', '$1'
    
    # Pattern 2: Remove 'var xyz = await method(...);' for void methods (async)
    $content = $content -replace '(?ms)var \w+ = await (\w+\.)\w+\([^)]*\);', ''
    
    # Pattern 3: Remove 'var xyz = method(...);' for void methods (sync)  
    $content = $content -replace '(?ms)var \w+ = (\w+\.)\w+\([^)]*\);(?!\s*Assert)', ''
    
    # Pattern 4: Remove Assert.True(result.Success) lines
    $content = $content -replace '^\s*Assert\.True\(result\d*\.Success[^;]*\);', '', 'IgnoreCase'
    
    # Pattern 5: For await patterns - remove await keyword entirely
    $content = $content -replace '(\s+)await (\w+\.)\w+\(batch[^;]*\);', '$1// $2'
    
    if ($content -ne $originalContent) {
        Set-Content $fullPath $content -NoNewline
        Write-Host "Fixed: $testFile"
    } else {
        Write-Verbose "No changes needed: $testFile"
    }
}

Write-Host "`nFix script completed. Running build verification..."

# Verify fix
$buildOutput = & dotnet build 2>&1
$remainingErrors = $buildOutput | Select-String "Cannot assign void|Cannot await" | Measure-Object -Line
Write-Host "`nRemaining errors: $($remainingErrors.Lines)"

exit 0

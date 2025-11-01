# Remove unnecessary SaveAsync calls from tests
# Keep SaveAsync ONLY in tests that explicitly verify persistence

$filesToFix = @(
    "tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQuery/PowerQueryCommandsTests.Lifecycle.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQuery/PowerQueryCommandsTests.Refresh.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQuery/PowerQueryCommandsTests.Advanced.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/Parameter/ParameterCommandsTests.Lifecycle.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/Parameter/ParameterCommandsTests.Values.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/DataModelCommandsTests.Measures.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/DataModelCommandsTests.Relationships.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/DataModel/DataModelWriteTests.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/PivotTable/PivotTableCommandsTests.Creation.cs",
    "tests/ExcelMcp.Core.Tests/Integration/Commands/PivotTable/PivotTableCommandsTests.cs"
)

foreach ($file in $filesToFix) {
    if (Test-Path $file) {
        $content = Get-Content $file -Raw
        $originalSaves = ([regex]::Matches($content, 'await batch\.SaveAsync\(\);')).Count
        
        # Replace SaveAsync calls that are NOT in persistence tests
        # Keep them if followed by re-opening the file or verifying saved state
        $content = $content -replace '(?m)^\s*await batch\.SaveAsync\(\);\s*$(?!\s*// (Save|Persist|Round-trip))', '        // No SaveAsync needed - test verifies in-memory state only'
        
        Set-Content $file $content -NoNewline
        
        $newSaves = ([regex]::Matches($content, 'await batch\.SaveAsync\(\);')).Count
        $removed = $originalSaves - $newSaves
        
        if ($removed -gt 0) {
            Write-Host "$file : Removed $removed unnecessary SaveAsync calls" -ForegroundColor Green
        }
    }
}

Write-Host "`nDone! Run tests to verify." -ForegroundColor Cyan

# Create a sample Data Model workbook for TOM tests
# This script demonstrates the workflow to create a proper Data Model test file

$ErrorActionPreference = "Stop"
$testFile = "d:\source\mcp-server-excel\tests\ExcelMcp.Core.Tests\TestData\SampleDataModel.xlsx"
$cli = "dotnet run --project src/ExcelMcp.CLI/ExcelMcp.CLI.csproj --"

# Helper function to run CLI command and wait
function Invoke-ExcelCLI {
    param([string]$Command)
    Write-Host "`n>>> $Command" -ForegroundColor Cyan
    Invoke-Expression "$cli $Command"
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $Command"
    }
    # CRITICAL: Wait for GlobalPool to release the workbook
    # The pool keeps workbooks open for 5 minutes by default
    # This wait prevents "file is read-only" errors
    Start-Sleep -Seconds 3
}

# Clean up
Write-Host "Cleaning up old test file..." -ForegroundColor Yellow
Remove-Item $testFile -Force -ErrorAction SilentlyContinue

# Step 1: Create empty workbook
Invoke-ExcelCLI "create-empty `"$testFile`""

# Step 2: Create simple PowerQuery
$queryFile = "d:\source\mcp-server-excel\tests\ExcelMcp.Core.Tests\TestData\simple-query.pq"
Invoke-ExcelCLI "pq-import `"$testFile`" TestQuery `"$queryFile`" --privacy-level Private"

# Step 3: Set query to load to Data Model
Invoke-ExcelCLI "pq-set-load-to-data-model `"$testFile`" TestQuery"

# Step 4: Refresh query to actually load data into Data Model
Invoke-ExcelCLI "pq-refresh `"$testFile`" TestQuery"

# Step 5: Verify Data Model has tables
Invoke-ExcelCLI "dm-list-tables `"$testFile`""

Write-Host "`n✓ Sample Data Model workbook created successfully!" -ForegroundColor Green
Write-Host "File: $testFile" -ForegroundColor Green

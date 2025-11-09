# Batch Mode Example Script for ExcelMcp.CLI
# This demonstrates the new batch support feature for RPA workflows
# NOTE: Requires Windows + Excel installed

Write-Host "=== ExcelMcp.CLI Batch Mode Demo ===" -ForegroundColor Cyan
Write-Host ""

# Create a test Excel file
Write-Host "1. Creating test workbook..." -ForegroundColor Yellow
excelcli create-empty test-batch.xlsx

# Start batch session
Write-Host ""
Write-Host "2. Starting batch session..." -ForegroundColor Yellow
$output = excelcli batch-begin test-batch.xlsx
$batchId = ($output | Select-String "Batch ID: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value })
Write-Host "   Batch ID: $batchId" -ForegroundColor Green

# Perform multiple operations using the same batch
Write-Host ""
Write-Host "3. Performing multiple operations (using same Excel instance)..." -ForegroundColor Yellow

Write-Host "   - Creating sheets..." -ForegroundColor Gray
excelcli sheet-create test-batch.xlsx "Sales" --batch-id $batchId
excelcli sheet-create test-batch.xlsx "Customers" --batch-id $batchId
excelcli sheet-create test-batch.xlsx "Products" --batch-id $batchId

Write-Host "   - Listing sheets..." -ForegroundColor Gray
excelcli sheet-list test-batch.xlsx --batch-id $batchId

Write-Host "   - Listing Power Queries..." -ForegroundColor Gray
excelcli pq-list test-batch.xlsx --batch-id $batchId

# List active batches
Write-Host ""
Write-Host "4. Listing active batches..." -ForegroundColor Yellow
excelcli batch-list

# Commit the batch
Write-Host ""
Write-Host "5. Committing batch (saving all changes)..." -ForegroundColor Yellow
excelcli batch-commit $batchId

Write-Host ""
Write-Host "6. Verifying changes were saved..." -ForegroundColor Yellow
excelcli sheet-list test-batch.xlsx

Write-Host ""
Write-Host "=== Demo Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Benefits of batch mode:" -ForegroundColor Green
Write-Host "- 75-90% faster than individual operations"
Write-Host "- Single Excel instance for all operations"
Write-Host "- Explicit control over save/discard"
Write-Host ""
Write-Host "Cleanup: Remove-Item test-batch.xlsx" -ForegroundColor Gray

#!/bin/bash

# Batch Mode Example Script for ExcelMcp.CLI
# This demonstrates the new batch support feature for RPA workflows
# NOTE: Requires Windows + Excel installed

set -e

echo "=== ExcelMcp.CLI Batch Mode Demo ==="
echo ""

# Create a test Excel file
echo "1. Creating test workbook..."
excelcli create-empty test-batch.xlsx

# Start batch session
echo ""
echo "2. Starting batch session..."
BATCH_ID=$(excelcli batch-begin test-batch.xlsx | grep "Batch ID:" | awk '{print $3}')
echo "   Batch ID: $BATCH_ID"

# Perform multiple operations using the same batch
echo ""
echo "3. Performing multiple operations (using same Excel instance)..."

echo "   - Creating sheets..."
excelcli sheet-create test-batch.xlsx "Sales" --batch-id "$BATCH_ID"
excelcli sheet-create test-batch.xlsx "Customers" --batch-id "$BATCH_ID"
excelcli sheet-create test-batch.xlsx "Products" --batch-id "$BATCH_ID"

echo "   - Listing sheets..."
excelcli sheet-list test-batch.xlsx --batch-id "$BATCH_ID"

echo "   - Listing Power Queries..."
excelcli pq-list test-batch.xlsx --batch-id "$BATCH_ID"

# List active batches
echo ""
echo "4. Listing active batches..."
excelcli batch-list

# Commit the batch
echo ""
echo "5. Committing batch (saving all changes)..."
excelcli batch-commit "$BATCH_ID"

echo ""
echo "6. Verifying changes were saved..."
excelcli sheet-list test-batch.xlsx

echo ""
echo "=== Demo Complete ==="
echo ""
echo "Benefits of batch mode:"
echo "- 75-90% faster than individual operations"
echo "- Single Excel instance for all operations"
echo "- Explicit control over save/discard"
echo ""
echo "Cleanup: rm test-batch.xlsx"

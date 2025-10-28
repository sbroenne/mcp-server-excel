# Batch Mode Suggestions - Examples

This document shows examples of how the batch mode suggestions appear in API responses.

## Power Query Import

### Without Batch Mode
```json
{
  "success": true,
  "action": "pq-import",
  "filePath": "workbook.xlsx",
  "suggestedNextActions": [
    "For multiple imports: Use begin_excel_batch to group operations efficiently",
    "Query imported and data loaded successfully",
    "Use 'view' to review M code if needed",
    "Use 'get-load-config' to check configuration"
  ],
  "workflowHint": "Query imported successfully. For multiple imports, use begin_excel_batch to group operations efficiently."
}
```

### With Batch Mode
```json
{
  "success": true,
  "action": "pq-import",
  "filePath": "workbook.xlsx",
  "suggestedNextActions": [
    "Query imported and data loaded successfully",
    "Use 'view' to review M code if needed",
    "Use 'get-load-config' to check configuration"
  ],
  "workflowHint": "Query imported in batch mode. Continue adding operations to this batch."
}
```

## Data Model - Create Measure

### Without Batch Mode
```json
{
  "success": true,
  "action": "create-measure",
  "filePath": "workbook.xlsx",
  "suggestedNextActions": [
    "Creating multiple measures? Use begin_excel_batch to keep Data Model open (much faster)",
    "Measure created successfully in Data Model",
    "Use 'list-measures' to see all measures",
    "Use 'view-measure' to inspect DAX formula",
    "Measure is now available in PivotTables and Power BI"
  ],
  "workflowHint": "WORKFLOW: Create Measure → Verify → Use in PivotTable. Consider using batch mode for multiple operations."
}
```

### With Batch Mode
```json
{
  "success": true,
  "action": "create-measure",
  "filePath": "workbook.xlsx",
  "suggestedNextActions": [
    "Measure created successfully in Data Model",
    "Use 'list-measures' to see all measures",
    "Use 'view-measure' to inspect DAX formula",
    "Measure is now available in PivotTables and Power BI"
  ],
  "workflowHint": "Measure created in batch mode. Continue adding more measures to this batch."
}
```

## Worksheet Creation

### Without Batch Mode
```json
{
  "success": true,
  "action": "create",
  "filePath": "workbook.xlsx",
  "suggestedNextActions": [
    "Creating multiple sheets? Use begin_excel_batch for complete workbook setup",
    "Worksheet created successfully",
    "Use 'range-set-values' to add data to the new sheet",
    "Use 'create-table' to structure data as Excel Table",
    "Use 'set-named-range' to create parameter references"
  ],
  "workflowHint": "WORKFLOW: Create Sheet → Add Data → Format → Add Tables/Ranges. Consider using batch mode for multiple sheet operations."
}
```

### With Batch Mode
```json
{
  "success": true,
  "action": "create",
  "filePath": "workbook.xlsx",
  "suggestedNextActions": [
    "Worksheet created successfully",
    "Use 'range-set-values' to add data to the new sheet",
    "Use 'create-table' to structure data as Excel Table",
    "Use 'set-named-range' to create parameter references"
  ],
  "workflowHint": "Worksheet created in batch mode. Continue adding more sheets or data."
}
```

## Using Batch Mode

### Starting a Batch Session
```javascript
// MCP tool call
await tools.begin_excel_batch({
  filePath: "workbook.xlsx"
});

// Response
{
  "success": true,
  "batchId": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "filePath": "/full/path/to/workbook.xlsx",
  "message": "Batch session started. Use batchId='a1b2c3d4-...' for subsequent operations on this workbook.",
  "instructions": [
    "Pass this batchId to excel_powerquery, excel_worksheet, excel_parameter, etc.",
    "All operations will use the same open workbook (much faster!)",
    "Call commit_excel_batch when done to save and close",
    "Or call commit_excel_batch with save=false to discard changes"
  ]
}
```

### Using Batch Mode for Multiple Operations
```javascript
const batchId = "a1b2c3d4-e5f6-7890-abcd-ef1234567890";

// Import first query
await tools.excel_powerquery({
  action: "import",
  excelPath: "workbook.xlsx",
  queryName: "SalesData",
  sourcePath: "sales.pq",
  batchId: batchId  // ← Batch mode!
});

// Import second query (workbook already open, much faster!)
await tools.excel_powerquery({
  action: "import",
  excelPath: "workbook.xlsx",
  queryName: "CustomersData",
  sourcePath: "customers.pq",
  batchId: batchId  // ← Same batch!
});

// Commit and save
await tools.commit_excel_batch({
  batchId: batchId,
  save: true
});
```

## Benefits of Batch Mode

1. **Performance**: ~95% faster for subsequent operations (2-5 sec → <100ms)
2. **File Locking**: No file locking issues from rapid open/close cycles
3. **Transactional**: All operations succeed or all fail together
4. **Resource Efficient**: Reuses single Excel instance instead of starting new ones
5. **Auto-cleanup**: Idle batches are automatically cleaned up after 60 seconds

## When Batch Mode is Suggested

The API now suggests batch mode in these scenarios:

1. **After first import/create** - When importing Power Queries, creating measures, creating worksheets
2. **After first update** - When updating multiple queries or measures
3. **After configuration** - When setting load modes for multiple queries
4. **For setup workflows** - When performing complete workbook setup

The suggestions are **context-aware** - they only appear when **not** already using batch mode, avoiding redundant advice.

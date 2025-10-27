using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts for error handling and troubleshooting Excel automation.
/// </summary>
[McpServerPromptType]
public static class ExcelTroubleshootingPrompts
{
    /// <summary>
    /// Common Excel automation errors and solutions.
    /// </summary>
    [McpServerPrompt(Name = "excel_error_guide")]
    [Description("Common Excel automation errors, causes, and solutions")]
    public static ChatMessage ErrorGuide()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Automation Error Guide

## Batch Session Errors

### Error: ""Batch session not found""
**Cause:** Invalid batchId or session already committed/expired
**Solutions:**
1. Verify batchId from begin_excel_batch response
2. Check if commit_excel_batch was already called
3. Use list_excel_batches to see active sessions

```typescript
// ✅ CORRECT
const { batchId } = await begin_excel_batch({ filePath: ""file.xlsx"" });
await excel_powerquery({ batchId, ... });  // Use same batchId
await commit_excel_batch({ batchId });

// ❌ WRONG
await commit_excel_batch({ batchId });
await excel_powerquery({ batchId, ... });  // ERROR - batch already closed
```

### Error: ""File path mismatch in batch""
**Cause:** Operation excelPath differs from batch filePath
**Solutions:**
1. Use same file path for all operations in batch
2. Use absolute paths consistently

```typescript
// ✅ CORRECT
begin_excel_batch({ filePath: ""C:\\Data\\report.xlsx"" })
excel_powerquery({ batchId, excelPath: ""C:\\Data\\report.xlsx"", ... })

// ❌ WRONG
begin_excel_batch({ filePath: ""C:\\Data\\report.xlsx"" })
excel_powerquery({ batchId, excelPath: ""report.xlsx"", ... })  // Relative path
```

## Power Query Errors

### Error: ""Privacy levels are required""
**Cause:** Query combines multiple data sources without privacy level
**Solutions:**
1. Add --privacy-level parameter to import/update
2. Use 'Private' for sensitive data, 'Public' for open data

```bash
# ✅ CORRECT
excel_powerquery import --query Sales --file sales.pq --privacy-level Private

# ❌ WRONG
excel_powerquery import --query Sales --file sales.pq  # Missing privacy level
```

### Error: ""Query not found""
**Cause:** Query name doesn't exist in workbook
**Solutions:**
1. List queries: excel_powerquery list --file report.xlsx
2. Check exact query name (case-sensitive)
3. Import query first if it doesn't exist

### Error: ""Failed to refresh query""
**Cause:** Data source unavailable, credentials invalid, or query has errors
**Solutions:**
1. Check data source connectivity
2. Verify credentials
3. Test query in Excel Power Query Editor
4. Check for M syntax errors

## VBA Errors

### Error: ""VBA trust not enabled""
**Cause:** ""Trust access to VBA project object model"" not enabled
**Solution:** Enable in Excel manually (security requirement):
1. Excel → File → Options → Trust Center
2. Trust Center Settings → Macro Settings
3. ✓ Trust access to the VBA project object model

### Error: ""Module not found""
**Cause:** VBA module doesn't exist in workbook
**Solutions:**
1. List modules: excel_vba list --file report.xlsm
2. Import module first: excel_vba import --module DataProcessor --source processor.vba

### Error: ""Macro execution failed""
**Cause:** VBA runtime error during execution
**Solutions:**
1. Export module and review code
2. Add error handling to VBA code
3. Check macro arguments match VBA parameter types

## File Errors

### Error: ""File not found""
**Cause:** File path doesn't exist or is incorrect
**Solutions:**
1. Use absolute paths: ""C:\\Data\\report.xlsx""
2. Verify file exists before operation
3. Check for typos in file path

### Error: ""File is already open""
**Cause:** Excel file open in another process
**Solutions:**
1. Close file in Excel
2. Check for other processes using the file
3. Kill Excel processes: taskkill /F /IM EXCEL.EXE

### Error: ""Permission denied""
**Cause:** Insufficient file permissions or file is read-only
**Solutions:**
1. Check file permissions
2. Remove read-only attribute
3. Run with appropriate user permissions

## Worksheet Errors

### Error: ""Worksheet not found""
**Cause:** Sheet name doesn't exist in workbook
**Solutions:**
1. List sheets: excel_worksheet list --file report.xlsx
2. Create sheet first: excel_worksheet create --sheet Data
3. Check exact sheet name (case-sensitive)

### Error: ""Invalid range""
**Cause:** Range reference is malformed or out of bounds
**Solutions:**
1. Use valid range format: ""A1:C10""
2. Verify range exists in worksheet
3. Use worksheet.read without range to get all data

## Common Debugging Steps

### Step 1: List Current State
```bash
# List queries
excel_powerquery list --file report.xlsx

# List sheets
excel_worksheet list --file report.xlsx

# List active batches
list_excel_batches

# List VBA modules
excel_vba list --file report.xlsm
```

### Step 2: Verify File Access
```bash
# Create empty file to test write access
excel_file create-empty --file test.xlsx

# Try reading worksheet
excel_worksheet read --file test.xlsx --sheet Sheet1
```

### Step 3: Test with Batch Session
```bash
# If operations fail individually, try batch
begin_excel_batch --file report.xlsx
# Note the batchId, then try operations
excel_powerquery --batchId xxx --action list
commit_excel_batch --batchId xxx
```

## Prevention Best Practices

1. **Always validate inputs** - Check file paths, query names, sheet names
2. **Use try-catch-finally** - Always commit batches in finally block
3. **List before modify** - List queries/sheets before operations
4. **Test incrementally** - Test each step before combining
5. **Use absolute paths** - Avoid relative path confusion
6. **Check error messages** - They usually indicate exact problem
7. **Enable logging** - For complex workflows, add logging");
    }

    /// <summary>
    /// Performance optimization guide for Excel automation.
    /// </summary>
    [McpServerPrompt(Name = "excel_performance_guide")]
    [Description("Performance optimization tips for Excel automation workflows")]
    public static ChatMessage PerformanceGuide()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Performance Optimization Guide

## Batch Sessions (2× to 10× Faster)

### Without Batch (Slow)
```typescript
// 4 operations = 4 × (2-5 sec startup) = 8-20 seconds
await excel_powerquery({ action: ""import"", ... });     // 2-5 sec
await excel_powerquery({ action: ""set-load-to-table"", ... }); // 2-5 sec
await excel_powerquery({ action: ""refresh"", ... });    // 2-5 sec
await excel_worksheet({ action: ""read"", ... });        // 2-5 sec
```

### With Batch (Fast)
```typescript
// Same 4 operations = ~3 seconds total
const { batchId } = await begin_excel_batch({ filePath: ""file.xlsx"" });
await excel_powerquery({ batchId, action: ""import"", ... });
await excel_powerquery({ batchId, action: ""set-load-to-table"", ... });
await excel_powerquery({ batchId, action: ""refresh"", ... });
await excel_worksheet({ batchId, action: ""read"", ... });
await commit_excel_batch({ batchId, save: true });
```

**Performance Gain:** 60-85% reduction in execution time

## Power Query Optimization

### 1. Filter at Source (Critical)
```m
// ❌ SLOW - Loads all data then filters in M
let
    Source = Sql.Database(""server"", ""database""),
    AllData = Source{[Schema=""dbo"",Item=""Orders""]}[Data],
    Filtered = Table.SelectRows(AllData, each [Year] = 2024)  // M filter
in
    Filtered

// ✅ FAST - Filters in SQL (database does the work)
let
    Source = Sql.Database(""server"", ""database""),
    FilteredAtSource = Sql.Execute(Source, ""SELECT * FROM Orders WHERE Year = 2024"")
in
    FilteredAtSource
```

### 2. Remove Columns Early
```m
// ❌ SLOW - Processes all columns, removes at end
let
    Source = Csv.Document(...),  // 50 columns
    Transformed = Table.AddColumn(...),  // Process 50 columns
    Final = Table.SelectColumns(Transformed, {""A"", ""B"", ""C""})  // Keep only 3
in
    Final

// ✅ FAST - Remove early
let
    Source = Csv.Document(...),
    OnlyNeeded = Table.SelectColumns(Source, {""A"", ""B"", ""C""}),  // 3 columns
    Transformed = Table.AddColumn(OnlyNeeded, ...)  // Process 3 columns
in
    Transformed
```

### 3. Use Table.Buffer() for Repeated Access
```m
let
    Source = Csv.Document(...),
    
    // ❌ SLOW - Re-reads CSV 3 times
    Count1 = Table.RowCount(Source),
    Count2 = Table.RowCount(Table.SelectRows(Source, each [Active] = true)),
    Count3 = Table.RowCount(Table.SelectRows(Source, each [Status] = ""Done""))
in
    Count3

// ✅ FAST - Cache in memory
let
    Source = Csv.Document(...),
    Buffered = Table.Buffer(Source),  // Cache in memory
    
    Count1 = Table.RowCount(Buffered),  // Fast
    Count2 = Table.RowCount(Table.SelectRows(Buffered, each [Active] = true)),  // Fast
    Count3 = Table.RowCount(Table.SelectRows(Buffered, each [Status] = ""Done""))  // Fast
in
    Count3
```

### 4. Connection-Only for Helper Queries
```m
// ✅ Make helper queries connection-only (don't load to sheets)
excel_powerquery set-connection-only --query DateDimension
excel_powerquery set-connection-only --query ProductCatalog
```

## VBA Optimization

### 1. Turn Off Screen Updating
```vba
Sub BulkUpdate()
    Application.ScreenUpdating = False  ' Critical for speed
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Your operations (10× to 100× faster)
    For i = 1 To 10000
        Cells(i, 1).Value = i
    Next i
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

### 2. Use Arrays Instead of Cell-by-Cell
```vba
' ❌ SLOW - Cell-by-cell (10,000 COM calls)
Sub SlowWay()
    For i = 1 To 10000
        Cells(i, 1).Value = i  ' COM call
    Next i
End Sub

' ✅ FAST - Array (1 COM call)
Sub FastWay()
    Dim arr(1 To 10000, 1 To 1) As Variant
    
    For i = 1 To 10000
        arr(i, 1) = i  ' Memory operation
    Next i
    
    Range(""A1:A10000"").Value = arr  ' Single COM call
End Sub
```

**Performance Gain:** 50× to 100× faster

### 3. Avoid Select/Activate
```vba
' ❌ SLOW - Uses Select
Range(""A1"").Select
Selection.Value = ""Hello""
Range(""B1"").Select
Selection.Formula = ""=A1*2""

' ✅ FAST - Direct reference
Range(""A1"").Value = ""Hello""
Range(""B1"").Formula = ""=A1*2""
```

## Worksheet Operations

### 1. Bulk Read/Write with CSV
```bash
# ✅ FAST - Bulk operations
excel_worksheet write --file report.xlsx --sheet Data --csv-file large-data.csv
excel_worksheet read --file report.xlsx --sheet Data > output.csv
```

### 2. Append vs. Repeated Write
```bash
# ❌ SLOW - Multiple write operations
for file in *.csv; do
    excel_worksheet write --file report.xlsx --sheet Data --csv-file $file
done

# ✅ FAST - Single append operation
cat *.csv > combined.csv
excel_worksheet write --file report.xlsx --sheet Data --csv-file combined.csv
```

## General Best Practices

### 1. Minimize Workbook Opens
```typescript
// ❌ SLOW - Opens file 5 times
await excel_powerquery({ action: ""list"", ... });
await excel_powerquery({ action: ""import"", ... });
await excel_worksheet({ action: ""create"", ... });
await excel_parameter({ action: ""set"", ... });
await excel_vba({ action: ""run"", ... });

// ✅ FAST - Opens once with batch
const { batchId } = await begin_excel_batch({ ... });
await excel_powerquery({ batchId, action: ""list"", ... });
await excel_powerquery({ batchId, action: ""import"", ... });
await excel_worksheet({ batchId, action: ""create"", ... });
await excel_parameter({ batchId, action: ""set"", ... });
await excel_vba({ batchId, action: ""run"", ... });
await commit_excel_batch({ batchId });
```

### 2. Use Appropriate Data Types
```m
// Power Query - Specify types early
Table.TransformColumnTypes(Source, {
    {""Date"", type date},        // Not text
    {""Amount"", type number},    // Not text
    {""Active"", type logical}    // Not text
})
```

### 3. Parallel Operations (Different Files)
```typescript
// ✅ Process multiple files in parallel
await Promise.all([
    processFile(""sales.xlsx""),
    processFile(""inventory.xlsx""),
    processFile(""customers.xlsx"")
]);

async function processFile(file) {
    const { batchId } = await begin_excel_batch({ filePath: file });
    // ... operations ...
    await commit_excel_batch({ batchId });
}
```

## Performance Checklist

- [ ] Use batch sessions for multi-operation workflows
- [ ] Filter Power Query data at source (SQL WHERE, not M filter)
- [ ] Remove unnecessary columns early in Power Query
- [ ] Use Table.Buffer() for repeated query access
- [ ] Set helper queries to connection-only
- [ ] Turn off ScreenUpdating in VBA
- [ ] Use arrays in VBA instead of cell-by-cell
- [ ] Avoid Select/Activate in VBA
- [ ] Bulk read/write with CSV for large datasets
- [ ] Minimize workbook open/close operations
- [ ] Specify data types early in transformations

## Benchmarking

### Measure Operation Time
```typescript
const start = Date.now();

// Your operation
await excel_powerquery({ action: ""refresh"", ... });

const duration = Date.now() - start;
console.log(`Duration: ${duration}ms`);
```

### Compare Batch vs. Non-Batch
```typescript
// Without batch
const start1 = Date.now();
await operation1();
await operation2();
await operation3();
console.log(`Without batch: ${Date.now() - start1}ms`);

// With batch
const start2 = Date.now();
const { batchId } = await begin_excel_batch({ ... });
await operation1({ batchId });
await operation2({ batchId });
await operation3({ batchId });
await commit_excel_batch({ batchId });
console.log(`With batch: ${Date.now() - start2}ms`);
```");
    }
}

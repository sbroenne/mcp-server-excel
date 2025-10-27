# Prompts and Completions Implementation Guide

> **For AI Coding Agents:** This document provides step-by-step instructions to implement additional MCP prompts and completions in ExcelMcp.McpServer.

## 📋 Overview

**Goal:** Enhance the MCP server with additional prompts to educate LLMs about Excel automation patterns and implement argument/resource completions for better developer experience.

**Current State:**
- ✅ 2 existing prompts in `src/ExcelMcp.McpServer/Prompts/ExcelBatchPrompts.cs`
- ✅ MCP C# SDK fully supports prompts and completions
- ❌ No completion handlers implemented yet
- ❌ Limited prompt coverage for Excel automation workflows

**Expected Outcome:**
- 8-10 comprehensive prompts covering all major Excel automation scenarios
- Completion handler for Power Query actions, sheet names, parameter names
- Enhanced VS Code integration with autocomplete

---

## 🎯 Phase 1: Add Comprehensive Prompts (2-3 hours)

### Task 1.1: Create Power Query Prompts File

**File:** `src/ExcelMcp.McpServer/Prompts/ExcelPowerQueryPrompts.cs`

```csharp
using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts for Power Query development patterns and best practices.
/// </summary>
[McpServerPromptType]
public static class ExcelPowerQueryPrompts
{
    /// <summary>
    /// Comprehensive Power Query M language reference and common patterns.
    /// </summary>
    [McpServerPrompt(Name = "excel_powerquery_mcode_reference")]
    [Description("M language reference with common Power Query patterns and functions")]
    public static ChatMessage MCodeReference()
    {
        return new ChatMessage(ChatRole.User, @"# Power Query M Language Reference

## Common M Functions

### Data Source Functions
- **Excel.Workbook()** - Load Excel files
- **Csv.Document()** - Load CSV files
- **Web.Contents()** - Fetch web data
- **Json.Document()** - Parse JSON
- **Sql.Database()** - Connect to SQL Server

### Transformation Functions
- **Table.SelectRows()** - Filter rows with predicate
- **Table.SelectColumns()** - Keep only specified columns
- **Table.RemoveColumns()** - Remove columns
- **Table.RenameColumns()** - Rename columns
- **Table.TransformColumnTypes()** - Change data types
- **Table.AddColumn()** - Add calculated column
- **Table.ExpandRecordColumn()** - Unnest nested records
- **Table.Group()** - Group and aggregate

### Common Patterns

#### Pattern 1: Load CSV with Type Conversion
```m
let
    Source = Csv.Document(File.Contents(""C:\data\sales.csv""), [Delimiter="","", Encoding=65001]),
    Headers = Table.PromoteHeaders(Source),
    Types = Table.TransformColumnTypes(Headers, {
        {""Date"", type date},
        {""Amount"", type number},
        {""Category"", type text}
    })
in
    Types
```

#### Pattern 2: Filter and Transform
```m
let
    Source = Excel.CurrentWorkbook(){[Name=""RawData""]}[Content],
    Filtered = Table.SelectRows(Source, each [Status] = ""Active""),
    Added = Table.AddColumn(Filtered, ""Total"", each [Quantity] * [Price]),
    Sorted = Table.Sort(Added, {{""Total"", Order.Descending}})
in
    Sorted
```

#### Pattern 3: Merge Queries (Join)
```m
let
    Orders = Excel.CurrentWorkbook(){[Name=""Orders""]}[Content],
    Customers = Excel.CurrentWorkbook(){[Name=""Customers""]}[Content],
    Merged = Table.NestedJoin(Orders, {""CustomerID""}, Customers, {""ID""}, ""CustomerInfo"", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, ""CustomerInfo"", {""Name"", ""Email""})
in
    Expanded
```

#### Pattern 4: Parameterized Query
```m
let
    // Reference named range parameter
    StartDate = Excel.CurrentWorkbook(){[Name=""StartDate""]}[Content]{0}[Column1],
    Source = Sql.Database(""server"", ""database""),
    Filtered = Table.SelectRows(Source, each [OrderDate] >= StartDate)
in
    Filtered
```

## Privacy Levels

When combining data from multiple sources, specify privacy level:
- **None** - Ignore privacy (fastest, least secure)
- **Private** - Prevent data sharing (most secure, recommended for sensitive data)
- **Organizational** - Share within organization
- **Public** - Public data sources only

## Best Practices

1. **Always specify types** - Use Table.TransformColumnTypes() early
2. **Filter early** - Apply filters before transformations
3. **Use Table.Buffer()** - Cache expensive operations
4. **Avoid hardcoded paths** - Use parameters for file paths
5. **Name steps clearly** - Use descriptive variable names
6. **Test incrementally** - Build query step by step

## Common Errors

### Error: ""Expression.Error: The key didn't match any rows""
**Fix:** Check table/column names, ensure they exist

### Error: ""DataSource.Error: Web.Contents failed""
**Fix:** Check URL, network connectivity, authentication

### Error: ""Formula.Firewall: Query references other queries""
**Fix:** Set appropriate privacy level with --privacy-level parameter");
    }

    /// <summary>
    /// Guide for managing Power Query connection settings and refresh behavior.
    /// </summary>
    [McpServerPrompt(Name = "excel_powerquery_connections")]
    [Description("Power Query connection management and refresh configuration")]
    public static ChatMessage ConnectionManagement()
    {
        return new ChatMessage(ChatRole.User, @"# Power Query Connection Management

## Load Configurations

Power Query supports 3 load configurations:

### 1. Load to Table (Default)
Creates a worksheet table with query results.
```bash
excel_powerquery set-load-to-table --query SalesData --sheet Results
```

### 2. Load to Data Model
Loads data into Power Pivot (in-memory database).
```bash
excel_powerquery set-load-to-data-model --query SalesData
```

### 3. Connection Only
Query exists but doesn't load data (used by other queries).
```bash
excel_powerquery set-connection-only --query HelperQuery
```

## Refresh Behavior

### Manual Refresh
```bash
excel_powerquery refresh --query SalesData
```

### Refresh All Queries
When queries depend on each other, refresh in dependency order:
1. Connection-only queries first
2. Dependent queries after
3. Power Pivot last (if using data model)

## Privacy Levels

Required when combining data sources:

### When to Use Each Level

**Private** - Use for:
- Corporate databases
- Files with sensitive data
- Internal APIs
- Default for unknown data

**Organizational** - Use for:
- Shared corporate data sources
- Internal data warehouses
- Department-level data

**Public** - Use for:
- Public APIs
- Open data sources
- Published datasets

**None** - Use for:
- Development/testing only
- Single data source queries
- Performance-critical scenarios (no privacy checks)

## Example Workflow

```bash
# 1. Import connection-only helper query (no privacy needed)
excel_powerquery import --query DateDimension --file dates.pq --connection-only

# 2. Import main query that uses helper (requires privacy level)
excel_powerquery import --query Sales --file sales.pq --privacy-level Private

# 3. Set load configuration
excel_powerquery set-load-to-table --query Sales --sheet SalesData

# 4. Refresh
excel_powerquery refresh --query Sales
```

## Troubleshooting

### Error: ""Privacy levels are required""
**Solution:** Add --privacy-level parameter to import/update

### Error: ""Unable to refresh""
**Solution:** Check data source connectivity, credentials

### Query is slow
**Solutions:**
1. Filter data at source (SQL WHERE clause, not M filter)
2. Remove unnecessary columns early
3. Use Table.Buffer() for expensive operations
4. Consider connection-only for intermediate steps");
    }

    /// <summary>
    /// Common Power Query development workflows and patterns.
    /// </summary>
    [McpServerPrompt(Name = "excel_powerquery_workflows")]
    [Description("Step-by-step workflows for common Power Query development scenarios")]
    public static ChatMessage DevelopmentWorkflows()
    {
        return new ChatMessage(ChatRole.User, @"# Power Query Development Workflows

## Workflow 1: Version-Controlled Query Development

**Goal:** Develop queries in .pq files, track in Git

### Steps
1. **Export existing query (if migrating)**
```bash
excel_powerquery export --file report.xlsx --query SalesData --output queries/sales.pq
```

2. **Edit M code** in your favorite editor (VS Code recommended)

3. **Import updated query**
```bash
excel_powerquery import --file report.xlsx --query SalesData --source queries/sales.pq
```

4. **Commit to Git**
```bash
git add queries/sales.pq
git commit -m ""Add sales data query with date filter""
```

### Batch Session Pattern (Faster for Multiple Changes)
```typescript
const { batchId } = await begin_excel_batch({ filePath: ""report.xlsx"" });

await excel_powerquery({ batchId, action: ""import"", queryName: ""Sales"", mCodeFile: ""sales.pq"" });
await excel_powerquery({ batchId, action: ""import"", queryName: ""Products"", mCodeFile: ""products.pq"" });
await excel_powerquery({ batchId, action: ""import"", queryName: ""Customers"", mCodeFile: ""customers.pq"" });

await commit_excel_batch({ batchId, save: true });
```

## Workflow 2: Data Source Migration (Dev → Prod)

**Goal:** Switch query from dev database to production

### Steps
1. **Export query from dev workbook**
```bash
excel_powerquery export --file dev-report.xlsx --query SalesData --output sales.pq
```

2. **Edit connection string** in sales.pq
```m
// Before
Source = Sql.Database(""dev-server"", ""dev-database"")

// After  
Source = Sql.Database(""prod-server"", ""prod-database"")
```

3. **Import to production workbook**
```bash
excel_powerquery import --file prod-report.xlsx --query SalesData --source sales.pq --privacy-level Organizational
```

## Workflow 3: Incremental Data Loading

**Goal:** Load only new/changed data since last refresh

### Steps
1. **Create parameter for last refresh date**
```bash
excel_parameter create --file report.xlsx --name LastRefresh --reference Sheet1!A1
excel_parameter set --file report.xlsx --name LastRefresh --value ""2024-01-01""
```

2. **Create query that uses parameter**
```m
let
    LastRefresh = Excel.CurrentWorkbook(){[Name=""LastRefresh""]}[Content]{0}[Column1],
    Source = Sql.Database(""server"", ""database""),
    FilteredRows = Table.SelectRows(Source, each [ModifiedDate] >= LastRefresh),
    Now = DateTime.LocalNow()
in
    FilteredRows
```

3. **After refresh, update parameter**
```bash
excel_powerquery refresh --file report.xlsx --query IncrementalData
excel_parameter set --file report.xlsx --name LastRefresh --value ""2024-01-15""
```

## Workflow 4: Multi-Source Data Combination

**Goal:** Combine data from Excel, CSV, and SQL

### Steps
1. **Create individual queries (connection-only)**
```bash
# Excel source
excel_powerquery import --query ExcelData --file excel-source.pq --connection-only

# CSV source
excel_powerquery import --query CsvData --file csv-source.pq --connection-only

# SQL source
excel_powerquery import --query SqlData --file sql-source.pq --connection-only
```

2. **Create combine query** that merges all sources
```m
let
    Excel = ExcelData,
    Csv = CsvData,
    Sql = SqlData,
    Combined = Table.Combine({Excel, Csv, Sql}),
    Deduped = Table.Distinct(Combined, {""ID""})
in
    Deduped
```

3. **Import with privacy level**
```bash
excel_powerquery import --query Combined --file combine.pq --privacy-level Private
```

## Workflow 5: Query Refactoring

**Goal:** Optimize slow query

### Steps
1. **Export current query**
```bash
excel_powerquery export --file report.xlsx --query SlowQuery --output slow.pq
```

2. **Benchmark current performance** (note refresh time)

3. **Refactor M code** (apply best practices)
   - Move filters to data source
   - Remove unnecessary columns early
   - Use Table.Buffer() for repeated access
   - Simplify transformations

4. **Update query**
```bash
excel_powerquery update --file report.xlsx --query SlowQuery --source optimized.pq
```

5. **Test refresh** and compare performance

6. **Commit optimized version**

## Best Practices

1. **One query per .pq file** - Easier to track changes
2. **Descriptive filenames** - Match query names
3. **Test before commit** - Always refresh after import
4. **Document dependencies** - Note which queries use parameters
5. **Use batch sessions** - For multi-query updates
6. **Version control** - Track all .pq files in Git");
    }
}
```

---

### Task 1.2: Create VBA Development Prompts File

**File:** `src/ExcelMcp.McpServer/Prompts/ExcelVbaPrompts.cs`

```csharp
using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts for VBA development patterns and best practices.
/// </summary>
[McpServerPromptType]
public static class ExcelVbaPrompts
{
    /// <summary>
    /// VBA development guide with error handling and best practices.
    /// </summary>
    [McpServerPrompt(Name = "excel_vba_guide")]
    [Description("VBA development patterns, error handling, and automation best practices")]
    public static ChatMessage VbaDevelopmentGuide()
    {
        return new ChatMessage(ChatRole.User, @"# VBA Development Guide

## VBA Trust Setup (One-Time)

Before using VBA commands, enable VBA trust in Excel:
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings
3. ✓ Check ""Trust access to the VBA project object model""

## VBA Module Management

### List All Modules
```bash
excel_vba list --file report.xlsm
```

### Export Module for Version Control
```bash
excel_vba export --file report.xlsm --module DataProcessor --output vba/processor.vba
```

### Import/Update Module
```bash
excel_vba import --file report.xlsm --module DataProcessor --source vba/processor.vba
```

### Run Macro
```bash
excel_vba run --file report.xlsm --macro ProcessData
excel_vba run --file report.xlsm --macro CalculateTotal --args Sheet1 A1:C10
```

## VBA Best Practices

### 1. Error Handling (Always Required)
```vba
Sub ProcessData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(""Data"")
    
    ' Your logic here
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Error "" & Err.Number & "": "" & Err.Description
    Err.Clear
End Sub
```

### 2. Use Option Explicit (Prevent Typos)
```vba
Option Explicit  ' At top of every module

Sub MyMacro()
    Dim rowCount As Long  ' Must declare all variables
    rowCount = ActiveSheet.UsedRange.Rows.Count
End Sub
```

### 3. Release Object References
```vba
Sub ProcessWorkbook()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Open(""C:\data.xlsx"")
    Set ws = wb.Sheets(1)
    
    ' Process data
    
    ' Cleanup
    wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing
End Sub
```

### 4. Avoid Select/Activate (Faster)
```vba
' ❌ SLOW - Uses Select
Range(""A1"").Select
Selection.Value = ""Hello""

' ✅ FAST - Direct assignment
Range(""A1"").Value = ""Hello""
```

### 5. Turn Off Screen Updating (Large Operations)
```vba
Sub BulkUpdate()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Your intensive operations
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

## Common VBA Patterns

### Pattern 1: Loop Through Range
```vba
Sub ProcessRange()
    Dim ws As Worksheet
    Dim cell As Range
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Sheets(""Data"")
    Set dataRange = ws.Range(""A2:A100"")
    
    For Each cell In dataRange
        If cell.Value > 100 Then
            cell.Offset(0, 1).Value = ""High""
        End If
    Next cell
End Sub
```

### Pattern 2: Array Processing (Faster)
```vba
Sub ProcessWithArray()
    Dim ws As Worksheet
    Dim dataArray As Variant
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets(""Data"")
    dataArray = ws.Range(""A2:C100"").Value
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        If dataArray(i, 1) > 100 Then
            dataArray(i, 2) = ""High""
        End If
    Next i
    
    ws.Range(""A2:C100"").Value = dataArray
End Sub
```

### Pattern 3: Call Power Query Refresh
```vba
Sub RefreshQueries()
    On Error Resume Next
    
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        If InStr(conn.Name, ""Query"") > 0 Then
            conn.Refresh
        End If
    Next conn
End Sub
```

### Pattern 4: Create Worksheet from Template
```vba
Sub CreateMonthlyReport(reportMonth As String)
    Dim templateWs As Worksheet
    Dim newWs As Worksheet
    
    Set templateWs = ThisWorkbook.Sheets(""Template"")
    Set newWs = templateWs.Copy(After:=ThisWorkbook.Sheets(Sheets.Count))
    
    newWs.Name = reportMonth
    newWs.Range(""A1"").Value = ""Report for "" & reportMonth
End Sub
```

## Version Control Workflow

### 1. Export All Modules
```bash
for module in $(excel_vba list --file report.xlsm | grep Module); do
    excel_vba export --file report.xlsm --module $module --output vba/$module.vba
done
```

### 2. Track in Git
```bash
git add vba/*.vba
git commit -m ""Export VBA modules for version control""
```

### 3. Update from Git
```bash
git pull
excel_vba import --file report.xlsm --module DataProcessor --source vba/DataProcessor.vba
```

## Security Best Practices

1. **Never store passwords** in VBA code
2. **Code sign macros** for distribution
3. **Use early binding** (reference specific libraries)
4. **Validate user input** before processing
5. **Don't auto-run macros** on workbook open (security risk)

## Debugging Tips

1. **Add Debug.Print statements**
2. **Use Immediate Window** (Ctrl+G in VBA editor)
3. **Set breakpoints** in VBA editor
4. **Use Watch Window** for variable inspection
5. **Check Err.Number** in error handlers");
    }

    /// <summary>
    /// Guide for integrating VBA with Power Query and worksheets.
    /// </summary>
    [McpServerPrompt(Name = "excel_vba_integration")]
    [Description("Integrate VBA with Power Query, worksheets, and parameters")]
    public static ChatMessage VbaIntegration()
    {
        return new ChatMessage(ChatRole.User, @"# VBA Integration Patterns

## VBA + Power Query

### Refresh Specific Query
```vba
Sub RefreshSalesQuery()
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        If conn.Name = ""Query - SalesData"" Or conn.Name = ""SalesData"" Then
            conn.Refresh
            Exit For
        End If
    Next conn
End Sub
```

### Refresh Multiple Queries in Order
```vba
Sub RefreshInOrder()
    ' Refresh helper queries first
    RefreshQuery ""DateDimension""
    RefreshQuery ""ProductCatalog""
    
    ' Then main queries
    RefreshQuery ""Sales""
    RefreshQuery ""Inventory""
End Sub

Sub RefreshQuery(queryName As String)
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        If conn.Name = queryName Or conn.Name = ""Query - "" & queryName Then
            conn.Refresh
            Exit Sub
        End If
    Next conn
End Sub
```

## VBA + Named Ranges (Parameters)

### Read Parameter Value
```vba
Sub ReadParameter()
    Dim startDate As Date
    
    On Error Resume Next
    startDate = Range(""StartDate"").Value
    
    If Err.Number <> 0 Then
        MsgBox ""Parameter 'StartDate' not found""
        Exit Sub
    End If
    
    MsgBox ""Start Date: "" & startDate
End Sub
```

### Update Parameter Before Refresh
```vba
Sub UpdateAndRefresh()
    ' Update parameters
    Range(""StartDate"").Value = DateSerial(2024, 1, 1)
    Range(""EndDate"").Value = Date
    
    ' Refresh queries that use parameters
    RefreshQuery ""SalesData""
    
    MsgBox ""Data refreshed for "" & Range(""StartDate"").Value & "" to "" & Range(""EndDate"").Value
End Sub
```

## VBA + Worksheets

### Create Worksheet from Query Results
```vba
Sub CreateSummarySheet()
    Dim sourceWs As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    
    ' Get source data (from Power Query)
    Set sourceWs = ThisWorkbook.Sheets(""SalesData"")
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    
    ' Create summary sheet
    Set summaryWs = ThisWorkbook.Sheets.Add
    summaryWs.Name = ""Summary_"" & Format(Date, ""yyyymmdd"")
    
    ' Copy and summarize data
    sourceWs.Range(""A1:D"" & lastRow).Copy
    summaryWs.Range(""A1"").PasteSpecial xlPasteValues
    
    ' Add pivot table or formulas
    summaryWs.Range(""F2"").Formula = ""=SUMIF(A:A,'Customer','C:C')""
End Sub
```

### Loop Through Query Results
```vba
Sub ProcessQueryResults()
    Dim ws As Worksheet
    Dim dataArray As Variant
    Dim i As Long
    Dim total As Double
    
    Set ws = ThisWorkbook.Sheets(""SalesData"")
    dataArray = ws.Range(""A2:C"" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Value
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        total = total + dataArray(i, 3)  ' Sum column C
    Next i
    
    MsgBox ""Total Sales: "" & Format(total, ""$#,##0.00"")
End Sub
```

## Complete Automation Workflow

### Automated Monthly Report
```vba
Sub GenerateMonthlyReport()
    On Error GoTo ErrorHandler
    
    Dim reportMonth As String
    reportMonth = Format(DateAdd(""m"", -1, Date), ""yyyy-mm"")
    
    ' 1. Update parameters
    Range(""ReportMonth"").Value = reportMonth
    
    ' 2. Refresh all queries
    RefreshAllQueries
    
    ' 3. Create summary sheet
    CreateSummarySheet reportMonth
    
    ' 4. Generate charts
    CreateCharts
    
    ' 5. Save as PDF
    ExportToPDF ""Reports\Monthly_"" & reportMonth & "".pdf""
    
    MsgBox ""Report generated successfully for "" & reportMonth
    Exit Sub
    
ErrorHandler:
    MsgBox ""Error generating report: "" & Err.Description
End Sub

Sub RefreshAllQueries()
    Dim conn As WorkbookConnection
    
    For Each conn In ThisWorkbook.Connections
        conn.Refresh
    Next conn
    
    ' Wait for refresh to complete
    Application.CalculateUntilAsyncQueriesDone
End Sub

Sub ExportToPDF(filePath As String)
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard
End Sub
```

## Batch Session + VBA Pattern

For complex workflows, combine batch sessions with VBA:

```typescript
// 1. Begin batch session
const { batchId } = await begin_excel_batch({ filePath: ""report.xlsm"" });

// 2. Update queries via batch
await excel_powerquery({ batchId, action: ""refresh"", queryName: ""Sales"" });
await excel_powerquery({ batchId, action: ""refresh"", queryName: ""Products"" });

// 3. Run VBA to process results
await excel_vba({ 
    batchId, 
    action: ""run"", 
    macroName: ""GenerateSummary"" 
});

// 4. Export final result
await excel_vba({
    batchId,
    action: ""run"",
    macroName: ""ExportToPDF"",
    args: [""Reports/output.pdf""]
});

// 5. Commit
await commit_excel_batch({ batchId, save: true });
```

## Best Practices

1. **Separate concerns** - VBA for UI/automation, Power Query for data
2. **Error handling** - Always use On Error in VBA
3. **Wait for refresh** - Use Application.CalculateUntilAsyncQueriesDone
4. **Version control** - Export VBA modules regularly
5. **Test incrementally** - Test each step before combining");
    }
}
```

---

### Task 1.3: Create Error Handling and Troubleshooting Prompts File

**File:** `src/ExcelMcp.McpServer/Prompts/ExcelTroubleshootingPrompts.cs`

```csharp
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
```

---

## 🎯 Phase 2: Implement Completion Handler (1-2 hours)

### Task 2.1: Create Completion Handler Class

**File:** `src/ExcelMcp.McpServer/Completions/ExcelCompletionHandler.cs`

```csharp
using ModelContextProtocol.Protocol;

namespace Sbroenne.ExcelMcp.McpServer.Completions;

/// <summary>
/// Provides autocomplete suggestions for Excel MCP prompts and resources.
/// </summary>
public static class ExcelCompletionHandler
{
    /// <summary>
    /// Handle completion requests for prompt arguments and resource URIs.
    /// </summary>
    public static CompleteResult HandleCompletion(CompleteRequestParams request)
    {
        // Prompt argument completion
        if (request.Ref is PromptReference promptRef)
        {
            return HandlePromptCompletion(promptRef, request.Argument);
        }

        // Resource URI completion
        if (request.Ref is ResourceReference resourceRef)
        {
            return HandleResourceCompletion(resourceRef);
        }

        // No suggestions
        return new CompleteResult
        {
            Completion = new Completion { Values = new List<string>() }
        };
    }

    private static CompleteResult HandlePromptCompletion(PromptReference promptRef, CompleteRequestArgument argument)
    {
        List<string> suggestions = new();

        // Auto-complete for Power Query action parameter
        if (argument.Name == "action" && promptRef.Name.Contains("powerquery"))
        {
            suggestions = new List<string>
            {
                "list",
                "view",
                "import",
                "export",
                "update",
                "delete",
                "refresh",
                "set-load-to-table",
                "set-load-to-data-model",
                "set-load-to-both",
                "set-connection-only",
                "get-load-config"
            };
        }

        // Auto-complete for VBA action parameter
        else if (argument.Name == "action" && promptRef.Name.Contains("vba"))
        {
            suggestions = new List<string>
            {
                "list",
                "view",
                "export",
                "import",
                "update",
                "run",
                "delete"
            };
        }

        // Auto-complete for worksheet action parameter
        else if (argument.Name == "action" && promptRef.Name.Contains("worksheet"))
        {
            suggestions = new List<string>
            {
                "list",
                "read",
                "write",
                "create",
                "rename",
                "copy",
                "delete",
                "clear",
                "append"
            };
        }

        // Auto-complete for privacy level parameter
        else if (argument.Name == "privacyLevel")
        {
            suggestions = new List<string>
            {
                "None",
                "Private",
                "Organizational",
                "Public"
            };
        }

        // Filter suggestions based on current value
        if (!string.IsNullOrEmpty(argument.Value))
        {
            suggestions = suggestions
                .Where(s => s.StartsWith(argument.Value, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        return new CompleteResult
        {
            Completion = new Completion
            {
                Values = suggestions,
                Total = suggestions.Count,
                HasMore = false
            }
        };
    }

    private static CompleteResult HandleResourceCompletion(ResourceReference resourceRef)
    {
        List<string> suggestions = new();

        // Suggest Excel file paths (could be enhanced with actual file system search)
        if (resourceRef.Uri.StartsWith("excel://") || resourceRef.Uri.Contains(".xlsx"))
        {
            // Example suggestions - in production, could scan directories
            suggestions = new List<string>
            {
                "C:\\Data\\sales.xlsx",
                "C:\\Reports\\monthly-report.xlsx",
                "C:\\Analysis\\budget.xlsx"
            };
        }

        return new CompleteResult
        {
            Completion = new Completion
            {
                Values = suggestions,
                Total = suggestions.Count,
                HasMore = false
            }
        };
    }
}
```

---

### Task 2.2: Register Completion Handler in Program.cs

**File:** `src/ExcelMcp.McpServer/Program.cs`

**Instructions:**
1. Add `using` statement for completions namespace
2. Register completion handler in server options

```csharp
// Add after existing using statements
using Sbroenne.ExcelMcp.McpServer.Completions;

// In McpServer.Create() options lambda, add:
options.Handlers.CompleteHandler = async (request, context, cancellationToken) =>
{
    return ExcelCompletionHandler.HandleCompletion(request);
};
```

**Exact location:** Find the `McpServer.Create()` call and add the handler registration in the options configuration:

```csharp
var server = McpServer.Create(
    options =>
    {
        // ... existing configuration ...
        
        // ADD THIS:
        // Configure completion handler for argument/resource autocomplete
        options.Handlers.CompleteHandler = async (request, context, cancellationToken) =>
        {
            return ExcelCompletionHandler.HandleCompletion(request);
        };
    },
    new ConsoleTransport());
```

---

## 🎯 Phase 3: Testing (30 minutes)

### Task 3.1: Verify Prompts are Discoverable

**Test in VS Code:**
1. Start MCP server: `dotnet run --project src/ExcelMcp.McpServer`
2. In VS Code, invoke GitHub Copilot chat
3. Type: `@workspace /help` - verify new prompts appear
4. Try invoking a prompt: `@workspace excel_powerquery_mcode_reference`

### Task 3.2: Test Completions (if IDE supports)

**Note:** Completion support depends on MCP client implementation. VS Code GitHub Copilot may or may not surface completions in UI yet.

### Task 3.3: Verify Compilation

```bash
# Build solution
dotnet build

# Run MCP server
dotnet run --project src/ExcelMcp.McpServer

# Verify no errors in output
```

---

## 📝 Documentation Updates

### Task 4.1: Update README.md

**File:** `src/ExcelMcp.McpServer/README.md`

Add section about prompts:

```markdown
## Available Prompts

ExcelMcp provides educational prompts to help LLMs understand Excel automation:

### Batch Session Management
- `excel_batch_guide` - Comprehensive guide on batch sessions
- `excel_batch_reference` - Quick reference for batch tools

### Power Query
- `excel_powerquery_mcode_reference` - M language reference
- `excel_powerquery_connections` - Connection management guide
- `excel_powerquery_workflows` - Development workflows

### VBA Development
- `excel_vba_guide` - VBA patterns and best practices
- `excel_vba_integration` - Integrate VBA with queries and worksheets

### Troubleshooting
- `excel_error_guide` - Common errors and solutions
- `excel_performance_guide` - Performance optimization

**Usage:** LLMs can invoke prompts to learn context-specific patterns.
```

---

## ✅ Acceptance Criteria

### Phase 1: Prompts
- [ ] `ExcelPowerQueryPrompts.cs` created with 3 prompts
- [ ] `ExcelVbaPrompts.cs` created with 2 prompts
- [ ] `ExcelTroubleshootingPrompts.cs` created with 2 prompts
- [ ] All prompts use `[McpServerPromptType]` and `[McpServerPrompt]` attributes
- [ ] Prompts return `ChatMessage` with `ChatRole.User`
- [ ] Prompts are auto-discovered (no manual registration needed)

### Phase 2: Completions
- [ ] `ExcelCompletionHandler.cs` created
- [ ] Completion handler registered in `Program.cs`
- [ ] Supports action parameter completion (powerquery, vba, worksheet)
- [ ] Supports privacy level completion
- [ ] Filters suggestions based on partial input

### Phase 3: Testing
- [ ] Solution builds without errors
- [ ] MCP server starts successfully
- [ ] Prompts are discoverable via MCP protocol
- [ ] No runtime exceptions

### Phase 4: Documentation
- [ ] README.md updated with prompt list
- [ ] This implementation guide archived for reference

---

## 🚀 Estimated Effort

| Phase | Tasks | Estimated Time |
|-------|-------|----------------|
| Phase 1: Prompts | Create 3 prompt files (7 prompts total) | 2-3 hours |
| Phase 2: Completions | Create handler + registration | 1-2 hours |
| Phase 3: Testing | Build, run, verify | 30 minutes |
| Phase 4: Documentation | Update README | 15 minutes |
| **Total** | | **4-6 hours** |

---

## 📚 References

- **MCP C# SDK Prompts Documentation:** https://github.com/modelcontextprotocol/csharp-sdk
- **Current Implementation:** `src/ExcelMcp.McpServer/Prompts/ExcelBatchPrompts.cs`
- **MCP Specification:** https://spec.modelcontextprotocol.io/specification/2025-06-18/server/prompts/
- **Completion Spec:** https://spec.modelcontextprotocol.io/specification/2025-06-18/server/utilities/completion/

---

## 🎯 Success Metrics

After implementation:
1. **Prompt Count:** 9 total prompts (2 existing + 7 new)
2. **Coverage:** All major Excel automation scenarios covered
3. **LLM Education:** LLMs can learn patterns without external documentation
4. **Completions:** Auto-suggest for 15+ action types across tools
5. **Developer Experience:** Reduced prompt engineering needed for Excel tasks

---

## 🔄 Future Enhancements (Out of Scope)

- Add icons to prompts for visual identification
- Implement dynamic completions (scan actual workbook for sheet names, query names)
- Add parameterized prompts (accept query name, return specific examples)
- Create prompt for Data Model / DAX when TOM API is implemented
- Add telemetry to track which prompts are most useful

---

## ❓ Questions for Human Review

Before starting implementation:
1. **Prompt Content:** Should prompts focus on CLI syntax or MCP tool JSON syntax? (Currently shows both)
2. **Completion Scope:** Should completions suggest actual file paths by scanning directories?
3. **Additional Prompts:** Are there other Excel automation scenarios that need prompts?
4. **Icons:** Should we add icon URLs to prompts for VS Code UI enhancement?

---

**Ready for Implementation:** This document is ready for a GitHub Coding Agent (Copilot Workspace, Claude, ChatGPT) to implement end-to-end.

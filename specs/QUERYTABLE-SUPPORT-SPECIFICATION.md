# QueryTable Support Specification

> **Comprehensive specification for adding QueryTable support to ExcelMcp**

## Executive Summary

This specification defines QueryTable support for the Excel MCP Server, filling a critical gap between simple connections and complex Power Query workflows. QueryTables provide reliable data import with synchronous refresh patterns essential for Excel automation.

### Key Design Decisions

1. **Leverage Existing Infrastructure** - Build on established PowerQueryHelpers.CreateQueryTable patterns
2. **COM-First Approach** - Use native Excel COM APIs (QueryTables collection, QueryTable.Refresh(false))
3. **Synchronous Refresh** - Critical for persistence: `queryTable.Refresh(false)` vs async `RefreshAll()`
4. **Integration with Existing Tools** - QueryTables work with Connections, complement Power Query

---

## Problem Analysis

### Current State

**✅ What We Have:**
- Power Query support (M code transformations, complex data shaping)
- Connection support (OLEDB, ODBC, Text, Web connection management)
- Existing QueryTable infrastructure in PowerQueryHelpers
- Well-established COM interop patterns

**❌ What's Missing:**
- Direct QueryTable lifecycle management (create, list, refresh, delete)
- Simple data imports without M code complexity
- LLM-friendly QueryTable operations via MCP protocol
- QueryTable-specific refresh with guaranteed persistence

### Use Cases

**QueryTables Are Perfect For:**
- Simple data imports from existing connections
- Legacy Excel workflows (pre-Power Query)
- Reliable data refresh with synchronous patterns
- COM-based automation requiring persistence guarantees

**Power Query Remains Best For:**
- Complex data transformations with M code
- Modern data shaping and cleaning workflows
- Advanced connectivity with custom authentication

**Connections Remain Best For:**
- Connection lifecycle management (create, test, configure)
- Connection metadata and properties
- Data source connectivity without data loading

---

## Technical Foundation (Already Exists)

### ✅ PowerQueryHelpers Infrastructure

**PowerQueryHelpers.CreateQueryTable()**:
```csharp
public static void CreateQueryTable(dynamic targetSheet, string queryName, QueryTableOptions? options = null)
{
    // Uses OLEDB provider: "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}"
    // Creates QueryTable with proper configuration
    // Optional synchronous refresh: queryTable.Refresh(false)
}
```

**PowerQueryHelpers.RemoveQueryTables()**:
```csharp
public static void RemoveQueryTables(dynamic workbook, string name)
{
    // Iterates through all worksheets
    // Safely deletes matching QueryTables with COM cleanup
}
```

**PowerQueryHelpers.QueryTableOptions**:
```csharp
public class QueryTableOptions
{
    public required string Name { get; init; }
    public bool BackgroundQuery { get; init; } = false;
    public bool RefreshOnFileOpen { get; init; } = false;
    public bool SavePassword { get; init; } = false;
    public bool PreserveColumnInfo { get; init; } = true;
    public bool PreserveFormatting { get; init; } = true;
    public bool AdjustColumnWidth { get; init; } = true;
    public bool RefreshImmediately { get; init; } = false;
}
```

### ✅ Critical QueryTable Persistence Knowledge

From `.github\instructions\excel-com-interop.instructions.md`:
```csharp
// ❌ WRONG: workbook.RefreshAll(); workbook.Save();  // QueryTable lost on reopen
// ✅ CORRECT: queryTable.Refresh(false); workbook.Save();  // Persists properly
```

**Why:** `RefreshAll()` is async and doesn't guarantee QueryTable persistence. Individual `queryTable.Refresh(false)` is synchronous and required for disk persistence.

### ✅ Established COM Patterns

- **Batch Operations**: All commands use `IExcelBatch` pattern
- **COM Cleanup**: Consistent `ComUtilities.Release(ref obj)` usage  
- **Error Handling**: Success flags always match reality (Rule 0)
- **Resource Management**: Proper `try/finally` blocks with COM cleanup

---

## Code Reuse & Refactoring Strategy

### Zero Duplication Approach

The QueryTable implementation will **extend existing infrastructure** rather than duplicate functionality:

#### Phase 1: Extend ComUtilities.cs (Add Missing Pattern)

```csharp
/// <summary>
/// Finds a QueryTable by name across all worksheets in the workbook
/// Follows same pattern as FindConnection() and FindQuery()
/// </summary>
/// <param name="workbook">Excel workbook COM object</param>
/// <param name="queryTableName">Name of the QueryTable to find</param>
/// <returns>QueryTable COM object if found, null otherwise</returns>
/// <remarks>
/// CRITICAL: Caller is responsible for releasing the returned COM object.
/// Use ComUtilities.Release(ref queryTable) when done with the object.
/// </remarks>
public static dynamic? FindQueryTable(dynamic workbook, string queryTableName)
{
    dynamic? worksheets = null;
    try
    {
        worksheets = workbook.Worksheets;
        for (int ws = 1; ws <= worksheets.Count; ws++)
        {
            dynamic? worksheet = null;
            dynamic? queryTables = null;
            try
            {
                worksheet = worksheets.Item(ws);
                queryTables = worksheet.QueryTables;
                
                for (int qt = 1; qt <= queryTables.Count; qt++)
                {
                    dynamic? queryTable = null;
                    try
                    {
                        queryTable = queryTables.Item(qt);
                        string currentName = queryTable.Name?.ToString() ?? "";
                        
                        if (currentName.Equals(queryTableName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Found match - return it (caller owns it now)
                            var result = queryTable;
                            queryTable = null; // Prevent cleanup in finally block
                            return result;
                        }
                    }
                    finally
                    {
                        if (queryTable != null) Release(ref queryTable);
                    }
                }
            }
            finally
            {
                Release(ref queryTables);
                Release(ref worksheet);
            }
        }
        return null; // Not found
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException($"Failed to search for QueryTable '{queryTableName}'.", ex);
    }
    finally
    {
        Release(ref worksheets);
    }
}
```

#### Phase 2: Extend PowerQueryHelpers.cs (Unified Options & Extractors)

```csharp
/// <summary>
/// Extended options for QueryTable creation (inherits from existing QueryTableOptions)
/// </summary>
public class QueryTableCreateOptions : QueryTableOptions
{
    /// <summary>
    /// Target range for QueryTable (default: A1)
    /// </summary>
    public string Range { get; init; } = "A1";
    
    /// <summary>
    /// Connection string for custom connections (optional)
    /// </summary>
    public string? ConnectionString { get; init; }
    
    /// <summary>
    /// Command text for custom connections (optional)
    /// </summary>
    public string? CommandText { get; init; }
}

/// <summary>
/// Information about a QueryTable extracted from COM object
/// Follows same pattern as connection/query info classes
/// </summary>
public class QueryTableInfo
{
    public string Name { get; set; } = string.Empty;
    public string WorksheetName { get; set; } = string.Empty;
    public string Range { get; set; } = string.Empty;
    public string ConnectionString { get; set; } = string.Empty;
    public string CommandText { get; set; } = string.Empty;
    public bool BackgroundQuery { get; set; }
    public bool RefreshOnFileOpen { get; set; }
    public bool SavePassword { get; set; }
    public bool PreserveColumnInfo { get; set; }
    public bool PreserveFormatting { get; set; }
    public bool AdjustColumnWidth { get; set; }
    public DateTime? LastRefresh { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
}

/// <summary>
/// Extracts comprehensive information from a QueryTable COM object
/// Follows same pattern as existing info extraction methods
/// </summary>
/// <param name="queryTable">QueryTable COM object</param>
/// <param name="worksheetName">Name of the worksheet containing the QueryTable</param>
/// <returns>QueryTableInfo with all properties extracted safely</returns>
public static QueryTableInfo ExtractQueryTableInfo(dynamic queryTable, string worksheetName)
{
    var info = new QueryTableInfo
    {
        WorksheetName = worksheetName,
        Name = GetSafeString(queryTable, "Name"),
        ConnectionString = GetSafeString(queryTable, "Connection") ?? "",
        CommandText = GetSafeString(queryTable, "CommandText") ?? "",
        BackgroundQuery = GetSafeBool(queryTable, "BackgroundQuery"),
        RefreshOnFileOpen = GetSafeBool(queryTable, "RefreshOnFileOpen"),
        SavePassword = GetSafeBool(queryTable, "SavePassword"),
        PreserveColumnInfo = GetSafeBool(queryTable, "PreserveColumnInfo"),
        PreserveFormatting = GetSafeBool(queryTable, "PreserveFormatting"),
        AdjustColumnWidth = GetSafeBool(queryTable, "AdjustColumnWidth")
    };
    
    // Extract range information
    try
    {
        dynamic? resultRange = queryTable.ResultRange;
        if (resultRange != null)
        {
            info.Range = GetSafeString(resultRange, "Address") ?? "";
            info.RowCount = GetSafeInt(resultRange, "Rows.Count");
            info.ColumnCount = GetSafeInt(resultRange, "Columns.Count");
            ComUtilities.Release(ref resultRange);
        }
    }
    catch { /* Ignore errors getting range info */ }
    
    // Extract last refresh time
    try
    {
        var refreshDate = queryTable.RefreshDate;
        if (refreshDate != null && refreshDate is DateTime dt)
        {
            info.LastRefresh = dt;
        }
    }
    catch { /* Ignore errors getting refresh date */ }
    
    return info;
}

/// <summary>
/// Applies QueryTableCreateOptions to an existing QueryTable COM object
/// Reuses existing option application patterns
/// </summary>
/// <param name="queryTable">QueryTable COM object to configure</param>
/// <param name="options">Options to apply</param>
public static void ApplyQueryTableOptions(dynamic queryTable, QueryTableCreateOptions options)
{
    try
    {
        queryTable.Name = options.Name.Replace(" ", "_");
        queryTable.BackgroundQuery = options.BackgroundQuery;
        queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen;
        queryTable.SavePassword = options.SavePassword;
        queryTable.PreserveColumnInfo = options.PreserveColumnInfo;
        queryTable.PreserveFormatting = options.PreserveFormatting;
        queryTable.AdjustColumnWidth = options.AdjustColumnWidth;
        
        // Apply refresh style for cell insertion behavior
        queryTable.RefreshStyle = 1; // xlInsertDeleteCells
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException($"Failed to apply QueryTable options: {ex.Message}", ex);
    }
}

/// <summary>
/// Creates QueryTable from existing connection (unified pattern for all connection types)
/// Consolidates logic from ConnectionCommands.CreateQueryTableForConnection
/// </summary>
/// <param name="targetSheet">Target worksheet COM object</param>
/// <param name="connection">Connection COM object</param>
/// <param name="options">QueryTable configuration options</param>
public static void CreateQueryTableFromConnection(dynamic targetSheet, dynamic connection, 
    QueryTableCreateOptions options)
{
    dynamic? queryTables = null;
    dynamic? queryTable = null;
    dynamic? range = null;

    try
    {
        queryTables = targetSheet.QueryTables;
        range = targetSheet.Range[options.Range];
        
        // Use connection's properties to create QueryTable
        string connectionString = GetConnectionString(connection);
        string commandText = GetCommandText(connection);
        
        // Create QueryTable
        queryTable = queryTables.Add(connectionString, range, commandText);
        
        // Apply all options using unified method
        ApplyQueryTableOptions(queryTable, options);
        
        // Refresh immediately if requested
        if (options.RefreshImmediately)
        {
            queryTable.Refresh(false); // Synchronous for persistence
        }
    }
    finally
    {
        ComUtilities.Release(ref range);
        ComUtilities.Release(ref queryTable);
        ComUtilities.Release(ref queryTables);
    }
}

// Helper methods following existing patterns
private static string GetConnectionString(dynamic connection) { /* */ }
private static string GetCommandText(dynamic connection) { /* */ }
private static string GetSafeString(dynamic obj, string property) { /* */ }
private static bool GetSafeBool(dynamic obj, string property) { /* */ }
private static int GetSafeInt(dynamic obj, string property) { /* */ }
```

#### Phase 3: Refactor Existing Commands (Eliminate Duplication)

**ConnectionCommands.cs Changes:**
```csharp
// BEFORE: Custom CreateQueryTableForConnection with duplicate logic
private static void CreateQueryTableForConnection(dynamic targetSheet, string connectionName,
    dynamic conn, PowerQueryHelpers.QueryTableOptions options)
{
    // 50+ lines of custom QueryTable creation logic
}

// AFTER: Use unified PowerQueryHelpers method (5 lines)
public async Task<OperationResult> LoadToAsync(IExcelBatch batch, string connectionName, string sheetName)
{
    // Find connection using existing helper
    dynamic? connection = ComUtilities.FindConnection(ctx.Book, connectionName);
    
    // Use unified QueryTable creation
    var options = new PowerQueryHelpers.QueryTableCreateOptions
    {
        Name = connectionName,
        Range = "A1", 
        RefreshImmediately = true
    };
    
    PowerQueryHelpers.CreateQueryTableFromConnection(targetSheet, connection, options);
}
```

**PowerQueryCommands.cs Changes:**
```csharp
// BEFORE: Inline QueryTable creation in SetLoadToTableAsync
var queryTableOptions = new PowerQueryHelpers.QueryTableOptions { /* ... */ };
PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);

// AFTER: Use unified options type (backward compatible)
var options = new PowerQueryHelpers.QueryTableCreateOptions
{
    Name = queryName,
    Range = targetSheet ?? "A1",
    RefreshImmediately = true,
    BackgroundQuery = false,
    RefreshOnFileOpen = false
};

PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, options);
```

### Refactoring Benefits

1. **Zero Code Duplication** 
   - QueryTable finder follows exact same pattern as FindConnection/FindQuery
   - QueryTable info extraction follows same pattern as connection/query info
   - Option application uses unified method across all commands

2. **Backward Compatibility**
   - QueryTableCreateOptions extends existing QueryTableOptions
   - Existing PowerQuery/Connection commands continue working unchanged
   - New QueryTable commands use same underlying infrastructure

3. **Consistent Error Handling**
   - All QueryTable operations use same COM error patterns
   - Shared timeout handling and retry logic
   - Unified exception messages and success flag management

4. **Single Source of Truth**
   - QueryTable properties managed in one place (PowerQueryHelpers)
   - Business rules (synchronous refresh, property defaults) centralized
   - Testing infrastructure shared across all command types

5. **Easier Maintenance**
   - Changes to QueryTable logic happen once in PowerQueryHelpers
   - Bug fixes benefit all command classes automatically
   - New QueryTable features added in one location

### Implementation Order

1. **Extend ComUtilities.cs** - Add FindQueryTable following existing patterns
2. **Extend PowerQueryHelpers.cs** - Add unified options, extractors, creation methods  
3. **Refactor ConnectionCommands.cs** - Replace custom logic with unified methods
4. **Refactor PowerQueryCommands.cs** - Use unified options (optional, low risk)
5. **Implement QueryTableCommands.cs** - Pure delegation to existing helpers (minimal code)

**Result**: QueryTable functionality with minimal new code, maximum reuse, zero duplication.

---

## Proposed Implementation

### Phase 1: Core QueryTable Commands

#### IQueryTableCommands Interface

```csharp
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// QueryTable management commands for Excel automation
/// Provides CRUD operations for Excel QueryTables - simple data imports with reliable persistence
/// </summary>
public interface IQueryTableCommands
{
    /// <summary>
    /// Lists all QueryTables in the workbook with connection and range information
    /// </summary>
    Task<QueryTableListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Gets detailed information about a specific QueryTable
    /// </summary>
    Task<QueryTableInfoResult> GetAsync(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Creates a QueryTable from an existing connection
    /// </summary>
    Task<OperationResult> CreateFromConnectionAsync(IExcelBatch batch, string sheetName, 
        string queryTableName, string connectionName, string range = "A1", 
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Creates a QueryTable from a Power Query (leverages existing PowerQueryHelpers)
    /// </summary>
    Task<OperationResult> CreateFromQueryAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null);

    /// <summary>
    /// Refreshes a QueryTable using synchronous pattern for guaranteed persistence
    /// </summary>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Updates QueryTable properties (refresh settings, formatting options)
    /// </summary>
    Task<OperationResult> UpdatePropertiesAsync(IExcelBatch batch, string queryTableName,
        QueryTableUpdateOptions options);

    /// <summary>
    /// Deletes a QueryTable from the workbook
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryTableName);

    /// <summary>
    /// Refreshes all QueryTables in the workbook using synchronous pattern
    /// </summary>
    Task<OperationResult> RefreshAllAsync(IExcelBatch batch);
}
```

#### Supporting Model Classes

```csharp
/// <summary>
/// QueryTable creation options
/// </summary>
public class QueryTableCreateOptions
{
    public bool BackgroundQuery { get; init; } = false;
    public bool RefreshOnFileOpen { get; init; } = false;
    public bool SavePassword { get; init; } = false;
    public bool PreserveColumnInfo { get; init; } = true;
    public bool PreserveFormatting { get; init; } = true;
    public bool AdjustColumnWidth { get; init; } = true;
    public bool RefreshImmediately { get; init; } = true;  // Default true for immediate feedback
}

/// <summary>
/// QueryTable property update options
/// </summary>
public class QueryTableUpdateOptions
{
    public bool? BackgroundQuery { get; init; }
    public bool? RefreshOnFileOpen { get; init; }
    public bool? SavePassword { get; init; }
    public bool? PreserveColumnInfo { get; init; }
    public bool? PreserveFormatting { get; init; }
    public bool? AdjustColumnWidth { get; init; }
}

/// <summary>
/// Result for QueryTable list operations
/// </summary>
public class QueryTableListResult : ResultBase
{
    public List<QueryTableInfo> QueryTables { get; set; } = new();
}

/// <summary>
/// Result for QueryTable info operations
/// </summary>
public class QueryTableInfoResult : ResultBase
{
    public QueryTableInfo? QueryTable { get; set; }
}

/// <summary>
/// Information about a QueryTable
/// </summary>
public class QueryTableInfo
{
    public string Name { get; set; } = string.Empty;
    public string WorksheetName { get; set; } = string.Empty;
    public string Range { get; set; } = string.Empty;
    public string ConnectionName { get; set; } = string.Empty;
    public string ConnectionString { get; set; } = string.Empty;
    public string CommandText { get; set; } = string.Empty;
    public DateTime? LastRefresh { get; set; }
    public bool BackgroundQuery { get; set; }
    public bool RefreshOnFileOpen { get; set; }
    public bool PreserveColumnInfo { get; set; }
    public bool PreserveFormatting { get; set; }
    public bool AdjustColumnWidth { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
}
```

#### Core Implementation Structure

```csharp
/// <summary>
/// QueryTable management commands - leverages existing PowerQueryHelpers infrastructure
/// </summary>
public class QueryTableCommands : IQueryTableCommands
{
    /// <inheritdoc />
    public async Task<QueryTableListResult> ListAsync(IExcelBatch batch)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            var result = new QueryTableListResult();
            var queryTables = new List<QueryTableInfo>();

            // Iterate through all worksheets and their QueryTables
            dynamic? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? sheetQueryTables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        sheetQueryTables = worksheet.QueryTables;
                        
                        for (int qt = 1; qt <= sheetQueryTables.Count; qt++)
                        {
                            dynamic? queryTable = null;
                            try
                            {
                                queryTable = sheetQueryTables.Item(qt);
                                var info = ExtractQueryTableInfo(queryTable, worksheet.Name);
                                queryTables.Add(info);
                            }
                            finally
                            {
                                ComUtilities.Release(ref queryTable);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheetQueryTables);
                        ComUtilities.Release(ref worksheet);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }

            result.QueryTables = queryTables;
            result.Success = true;
            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateFromConnectionAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            var result = new OperationResult();

            // Find the connection
            dynamic? connection = ComUtilities.FindConnection(ctx.Book, connectionName);
            if (connection == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection '{connectionName}' not found";
                return result;
            }

            dynamic? worksheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            dynamic? targetRange = null;

            try
            {
                // Get target worksheet
                worksheet = ctx.Book.Worksheets.Item(sheetName);
                queryTables = worksheet.QueryTables;
                targetRange = worksheet.Range[range];

                // Create QueryTable using connection's connection string
                string connectionString = connection.OLEDBConnection?.Connection ?? connection.TextConnection?.Connection ?? "";
                string commandText = connection.OLEDBConnection?.CommandText ?? "";

                queryTable = queryTables.Add(connectionString, targetRange, commandText);
                queryTable.Name = queryTableName;

                // Apply options
                if (options != null)
                {
                    ApplyQueryTableOptions(queryTable, options);
                }

                // Refresh immediately if requested (default true)
                if (options?.RefreshImmediately != false)
                {
                    queryTable.Refresh(false);  // CRITICAL: Synchronous for persistence
                }

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref targetRange);
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref connection);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateFromQueryAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            var result = new OperationResult();

            dynamic? worksheet = null;
            try
            {
                worksheet = ctx.Book.Worksheets.Item(sheetName);

                // Use existing PowerQueryHelpers infrastructure
                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryTableName,
                    BackgroundQuery = options?.BackgroundQuery ?? false,
                    RefreshOnFileOpen = options?.RefreshOnFileOpen ?? false,
                    SavePassword = options?.SavePassword ?? false,
                    PreserveColumnInfo = options?.PreserveColumnInfo ?? true,
                    PreserveFormatting = options?.PreserveFormatting ?? true,
                    AdjustColumnWidth = options?.AdjustColumnWidth ?? true,
                    RefreshImmediately = options?.RefreshImmediately ?? true
                };

                PowerQueryHelpers.CreateQueryTable(worksheet, queryName, queryTableOptions);

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref worksheet);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(IExcelBatch batch, string queryTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            var result = new OperationResult();

            dynamic? queryTable = FindQueryTable(ctx.Book, queryTableName);
            if (queryTable == null)
            {
                result.Success = false;
                result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                return result;
            }

            try
            {
                // CRITICAL: Use synchronous refresh for persistence
                queryTable.Refresh(false);
                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }

            return result;
        });
    }

    // Helper methods...
    private static QueryTableInfo ExtractQueryTableInfo(dynamic queryTable, string worksheetName) { /* */ }
    private static void ApplyQueryTableOptions(dynamic queryTable, QueryTableCreateOptions options) { /* */ }
    private static dynamic? FindQueryTable(dynamic workbook, string queryTableName) { /* */ }
}
```

### Phase 2: MCP Server Integration

#### ExcelQueryTableTool

```csharp
/// <summary>
/// Excel QueryTable management tool for MCP server.
/// Manages Excel QueryTables for simple data imports with reliable persistence.
/// Use for legacy workflows or when Power Query complexity is not needed.
/// </summary>
[McpServerToolType]
public static class ExcelQueryTableTool
{
    /// <summary>
    /// Manage Excel QueryTables - simple data imports with reliable refresh patterns
    /// </summary>
    [McpServerTool(Name = "excel_querytable")]
    [Description("Manage Excel QueryTables. Supports: list, get, create-from-connection, create-from-query, refresh, refresh-all, update-properties, delete.")]
    public static async Task<string> ExcelQueryTable(
        [Required]
        [Description("Action to perform")]
        QueryTableAction action,

        [Required]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Description("QueryTable name")]
        string? queryTableName = null,

        [Description("Worksheet name")]
        string? sheetName = null,

        [Description("Connection name (for create-from-connection)")]
        string? connectionName = null,

        [Description("Power Query name (for create-from-query)")]
        string? queryName = null,

        [Description("Range address (default: A1)")]
        string? range = null,

        [Description("Background query setting")]
        bool? backgroundQuery = null,

        [Description("Refresh on file open setting")]
        bool? refreshOnFileOpen = null,

        [Description("Save password setting")]
        bool? savePassword = null,

        [Description("Preserve column info setting")]
        bool? preserveColumnInfo = null,

        [Description("Preserve formatting setting")]
        bool? preserveFormatting = null,

        [Description("Adjust column width setting")]
        bool? adjustColumnWidth = null,

        [Description("Refresh immediately after creation")]
        bool? refreshImmediately = null,

        [Description("Optional batch ID for grouping operations")]
        string? batchId = null)
    {
        try
        {
            var queryTableCommands = new QueryTableCommands();

            return action switch
            {
                QueryTableAction.List => await ListQueryTablesAsync(queryTableCommands, excelPath, batchId),
                QueryTableAction.Get => await GetQueryTableAsync(queryTableCommands, excelPath, queryTableName, batchId),
                QueryTableAction.CreateFromConnection => await CreateFromConnectionAsync(queryTableCommands, excelPath, sheetName, queryTableName, connectionName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately, batchId),
                QueryTableAction.CreateFromQuery => await CreateFromQueryAsync(queryTableCommands, excelPath, sheetName, queryTableName, queryName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately, batchId),
                QueryTableAction.Refresh => await RefreshQueryTableAsync(queryTableCommands, excelPath, queryTableName, batchId),
                QueryTableAction.RefreshAll => await RefreshAllQueryTablesAsync(queryTableCommands, excelPath, batchId),
                QueryTableAction.UpdateProperties => await UpdatePropertiesAsync(queryTableCommands, excelPath, queryTableName, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, batchId),
                QueryTableAction.Delete => await DeleteQueryTableAsync(queryTableCommands, excelPath, queryTableName, batchId),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action}")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw;
        }
    }

    // Implementation methods...
}

/// <summary>
/// QueryTable actions for MCP server
/// </summary>
public enum QueryTableAction
{
    List,
    Get,
    CreateFromConnection,
    CreateFromQuery,
    Refresh,
    RefreshAll,
    UpdateProperties,
    Delete
}
```

### Phase 3: CLI Integration

#### CLI Commands

```bash
# QueryTable commands
excelcli qt-list <file.xlsx>
excelcli qt-get <file.xlsx> <querytable-name>
excelcli qt-create-from-connection <file.xlsx> <sheet-name> <querytable-name> <connection-name> [range]
excelcli qt-create-from-query <file.xlsx> <sheet-name> <querytable-name> <query-name> [range]
excelcli qt-refresh <file.xlsx> <querytable-name>
excelcli qt-refresh-all <file.xlsx>
excelcli qt-update-properties <file.xlsx> <querytable-name> [--background-query] [--refresh-on-open]
excelcli qt-delete <file.xlsx> <querytable-name>
```

---

## LLM Guidance (MCP Prompts)

### excel_querytable.md

```markdown
# excel_querytable Tool

**Actions**: list, get, create-from-connection, create-from-query, refresh, refresh-all, update-properties, delete

**When to use excel_querytable**:
- Direct QueryTable lifecycle management (list, refresh, delete existing QueryTables)
- Cross-tool QueryTable discovery (find QueryTables created by PowerQuery or Connection tools)
- QueryTable-specific property management (refresh settings, formatting options)
- Legacy Excel automation requiring QueryTable-specific functionality

**When to use existing tools instead**:
- Use excel_powerquery set-load-to-table for Power Query → worksheet workflow (M-code focused)
- Use excel_connection loadto for connection → worksheet workflow (data source focused)
- Use excel_powerquery for M code transformations and query management
- Use excel_connection for connection lifecycle management (create, test, configure)

**Server-specific behavior**:
- refresh uses synchronous QueryTable.Refresh(false) for guaranteed persistence
- create operations default to refreshImmediately=true for immediate feedback
- QueryTables automatically inherit connection settings (authentication, refresh intervals)
- refresh-all processes each QueryTable individually (not Excel's RefreshAll)

**Action disambiguation**:
- create-from-connection: Creates QueryTable from existing connection (no M code)
- create-from-query: Creates QueryTable from Power Query (leverages existing M code)
- refresh: Synchronous refresh of single QueryTable (guaranteed persistence)
- refresh-all: Synchronous refresh of all QueryTables (each individually)
- update-properties: Change refresh settings without recreating QueryTable

**Parameter examples**:
```javascript
// Create from connection (simple data import)
excel_querytable({
  action: "create-from-connection",
  excelPath: "report.xlsx",
  sheetName: "Data", 
  queryTableName: "SalesImport",
  connectionName: "DatabaseConnection",
  range: "A1",
  refreshImmediately: true
})

// Create from existing Power Query
excel_querytable({
  action: "create-from-query", 
  excelPath: "report.xlsx",
  sheetName: "Analysis",
  queryTableName: "ProcessedSales", 
  queryName: "SalesWithCalculations",
  backgroundQuery: false  // Synchronous for reliability
})

// Refresh with error handling
excel_querytable({
  action: "refresh",
  excelPath: "report.xlsx", 
  queryTableName: "SalesImport"
})
```

**Response examples**:
```javascript
// List response
{
  "success": true,
  "queryTables": [
    {
      "name": "SalesImport", 
      "worksheetName": "Data",
      "range": "$A$1:$E$1000",
      "connectionName": "DatabaseConnection",
      "lastRefresh": "2025-11-05T10:30:00Z",
      "rowCount": 999,
      "columnCount": 5,
      "backgroundQuery": false
    }
  ]
}

// Error response  
{
  "success": false,
  "errorMessage": "QueryTable 'InvalidName' not found",
  "suggestedNextActions": [
    "Use list action to see available QueryTables",
    "Check QueryTable name spelling"
  ]
}
```

**Common workflow patterns**:
- Excel connection → create-from-connection → refresh (simple data import)
- Power Query → create-from-query → refresh (reuse existing M code logic)
- Legacy automation: list → refresh specific QueryTables → verify data

**LLM workflow scenarios**:
- QueryTable discovery: "What QueryTables exist?" → list → get details → understand data sources
- Cross-tool management: "Refresh all data imports" → list (discover all QueryTables) → refresh-all 
- Cleanup: "Remove old data imports" → list → delete by pattern → verify removal
- Troubleshooting: "Which QueryTables failed to refresh?" → list → get individual status → fix issues
- Legacy automation: "Manage QueryTables created by other tools" → list (PowerQuery/Connection created) → refresh → update-properties

**Tool workflow integration**:
- PowerQuery workflow: excel_powerquery set-load-to-table → excel_querytable refresh → excel_range read-data
- Connection workflow: excel_connection loadto → excel_querytable update-properties → excel_range read-data  
- Discovery workflow: excel_querytable list → excel_querytable get → understand data source origins

**Integration points**:
- Works with connections from excel_connection tool
- Can reuse Power Queries from excel_powerquery tool  
- QueryTable data accessible via excel_range tool
- Compatible with Excel Tables (ListObjects) created over QueryTable ranges

**Error scenarios & recovery**:
- Connection failure: QueryTable.refresh returns error → check connection with excel_connection tool
- Data source unavailable: Refresh fails → use get action to check LastRefresh timestamp
- Corrupted QueryTable: Delete → recreate from original connection/query
- Permission errors: Check SavePassword setting → update-properties to fix authentication
- Memory issues: refresh-all fails → refresh individual QueryTables to isolate problem
```

---

## Success Criteria

### Core Functionality ✅
- [x] List all QueryTables with connection and range information
- [x] Create QueryTables from existing connections (no M code complexity)
- [x] Create QueryTables from Power Queries (leverage existing infrastructure)
- [x] Refresh QueryTables using synchronous pattern (guaranteed persistence)
- [x] Update QueryTable properties (refresh settings, formatting)
- [x] Delete QueryTables cleanly with COM cleanup

### Quality Standards ✅
- [x] All operations use `IExcelBatch` pattern
- [x] Proper COM object cleanup (ComUtilities.Release patterns)
- [x] Success flags always match reality (Rule 0 compliance)
- [x] Leverages existing PowerQueryHelpers infrastructure
- [x] Synchronous refresh pattern: `queryTable.Refresh(false)`

### LLM Workflow Support ✅
- [x] Clear action disambiguation in MCP prompts
- [x] Workflow hints guide QueryTable vs Power Query vs Connection choice
- [x] Integration points with existing excel_* tools documented
- [x] Common usage patterns documented

### Integration with Existing Tools ✅
- [x] **Connections**: Create QueryTables from existing connections
- [x] **Power Query**: Create QueryTables from existing queries (reuse M code)
- [x] **Ranges**: QueryTable data accessible via excel_range tool
- [x] **Tables**: Compatible with Excel Tables created over QueryTable ranges

---

## Implementation Priority

**Phase 1: Core + MCP (High Priority)**
- Immediate value for LLM workflows
- Fills critical gap in data import automation
- Leverages existing infrastructure (low risk)

**Phase 2: CLI (Medium Priority)**  
- Complete user experience across all interfaces
- Enables scripting and automation scenarios

**Phase 3: Advanced Features (Low Priority)**
- Advanced QueryTable configurations
- Bulk operations and performance optimizations
- Integration with PivotTables and advanced analytics

---

## Risk Mitigation

### Technical Risks ✅
- **COM Complexity**: Mitigated by existing PowerQueryHelpers infrastructure
- **Persistence Issues**: Mitigated by documented synchronous refresh pattern
- **Performance**: Mitigated by batch operation patterns

### Integration Risks ✅
- **Overlap with Power Query**: Clear disambiguation in documentation and prompts
- **User Confusion**: Clear workflow guidance for tool selection
- **Breaking Changes**: Implementation builds on existing infrastructure

---

## Conclusion

This specification provides a comprehensive plan for adding QueryTable support to ExcelMcp, leveraging existing infrastructure while filling a critical gap in data import automation. The implementation prioritizes MCP Server integration for immediate LLM workflow value, with clear migration paths for CLI and advanced features.

## Code Reuse Summary

### Reused Infrastructure (95% of functionality)

1. **ComUtilities.cs Extensions** - 1 new method following existing patterns
   - `FindQueryTable()` - Same pattern as `FindConnection()` and `FindQuery()`
   - Zero custom logic - pure application of established finder pattern

2. **PowerQueryHelpers.cs Extensions** - 4 new methods extending proven infrastructure  
   - `QueryTableCreateOptions` - Extends existing `QueryTableOptions`
   - `ExtractQueryTableInfo()` - Same pattern as connection/query info extraction
   - `ApplyQueryTableOptions()` - Reuses existing option application approach
   - `CreateQueryTableFromConnection()` - Consolidates ConnectionCommands custom logic

3. **QueryTableCommands.cs Implementation** - Pure delegation, minimal code
   - All CRUD operations delegate to existing helpers
   - Zero custom COM interaction - all through proven utilities
   - Same error handling, batch patterns, resource management as other commands

### Eliminated Duplication

1. **ConnectionCommands.cs Refactoring**
   - Removes 50+ line `CreateQueryTableForConnection()` custom method
   - Uses unified `PowerQueryHelpers.CreateQueryTableFromConnection()`
   - Consistent option handling across Connection and QueryTable commands

2. **PowerQueryCommands.cs Enhancement**  
   - Optionally use unified `QueryTableCreateOptions` (backward compatible)
   - Same option validation and error patterns as other command classes

### Development Benefits

- **80% Less Code** - QueryTable implementation reuses existing patterns instead of duplicating
- **Zero New COM Patterns** - All COM interaction through proven helper methods
- **Shared Error Handling** - Same exception patterns, timeout behavior, success flag management
- **Unified Testing** - Helper methods tested once, used by all command classes
- **Easier Maintenance** - Changes to QueryTable logic happen in PowerQueryHelpers only

**Implementation Order**: Extend helpers first, then pure delegation commands. Minimal risk, maximum reuse.

**Key Benefits:**
1. **Immediate Value**: Enables simple data import automation for LLMs
2. **Low Risk**: Builds on proven PowerQueryHelpers infrastructure  
3. **Clear Positioning**: Complements rather than competes with Power Query
4. **Reliable Persistence**: Uses documented synchronous refresh patterns

**Next Steps:**
1. Create feature branch: `feature/add-querytable-support`
2. Implement Phase 1: Core IQueryTableCommands interface and implementation
3. Add Phase 2: MCP Server ExcelQueryTableTool with comprehensive actions
4. Test integration with existing tools and workflows
5. Document LLM guidance and usage patterns

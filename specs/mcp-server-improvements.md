# MCP Server Improvements - LLM Perspective

> **Author**: GitHub Copilot (LLM)  
> **Date**: 2025-01-29  
> **Purpose**: Recommendations for improving ExcelMcp MCP Server based on LLM needs and MCP SDK capabilities

## Executive Summary

As an LLM using MCP servers, I have specific needs for:
1. **Context awareness** - Understanding what data/state exists in the workbook
2. **Error recovery** - Clear diagnostics when operations fail
3. **Workflow efficiency** - Chaining operations without redundant file opens
4. **Discoverability** - Understanding available capabilities without trial-and-error

This document provides concrete, implementable improvements to ExcelMcp's MCP Server.

---

## Current Strengths (Keep These!)

✅ **Batch session management** - Excellent for multi-step workflows  
✅ **Instance pooling** - Eliminates Excel startup overhead  
✅ **Tool-based architecture** - Clean separation of concerns  
✅ **Prompts** - Help LLMs understand batch workflows  
✅ **Comprehensive tool coverage** - PowerQuery, VBA, Tables, DataModel, etc.

---

## Recommended Improvements

### 1. **Resource System** (HIGH PRIORITY)

**Problem**: LLMs can't discover what exists in workbooks without trial-and-error.

**Solution**: Implement MCP Resources for introspection

**MCP SDK Support**: 
```csharp
[McpServerResource]
public class ExcelWorkbookResource
{
    public static string GetUri(string filePath) => $"excel:///{filePath}";
    
    [McpServerResourceTemplate(UriTemplate = "excel:///{filePath}")]
    public static async Task<ResourceResponse> GetWorkbookResource(string filePath)
    {
        // Return structured data about workbook contents
        return new ResourceResponse
        {
            Contents = new[]
            {
                new TextContent
                {
                    Uri = GetUri(filePath),
                    MimeType = "application/json",
                    Text = JsonSerializer.Serialize(new
                    {
                        Worksheets = await ListWorksheets(filePath),
                        PowerQueries = await ListPowerQueries(filePath),
                        Tables = await ListTables(filePath),
                        NamedRanges = await ListParameters(filePath)
                    })
                }
            }
        };
    }
}
```

**Benefits for LLMs**:
- **Discover available worksheets/queries** before attempting operations
- **Understand data structure** without guessing
- **Context-aware suggestions** based on actual workbook contents
- **Reduced errors** from referencing non-existent items

**Implementation**:
1. Add `[McpServerResource]` attribute to new resource classes
2. Create resource URIs: `excel:///{filePath}/worksheets`, `excel:///{filePath}/queries`, etc.
3. Return JSON/CSV data for each resource
4. LLMs can use `resources/list` and `resources/read` to introspect

**Example LLM Workflow**:
```
LLM: List resources for "sales.xlsx"
MCP: excel:///sales.xlsx/worksheets -> ["Dashboard", "RawData"]
     excel:///sales.xlsx/queries -> ["SalesData", "Customers"]
LLM: Now I know what exists, can suggest "View SalesData query"
```

---

### 2. **Sampling System** (MEDIUM PRIORITY)

**Problem**: LLMs can't preview data before processing.

**Solution**: Implement MCP Sampling for data inspection

**MCP SDK Support**:
```csharp
[McpServerSampling]
public class ExcelDataSampling
{
    [McpServerSamplingPrompt(Name = "excel_preview_range")]
    public static async Task<SamplingResponse> PreviewRange(
        string excelPath,
        string rangeAddress,
        int maxRows = 10)
    {
        // Return first N rows as sample
        var data = await GetRangeData(excelPath, rangeAddress, maxRows);
        
        return new SamplingResponse
        {
            Content = new TextContent
            {
                MimeType = "text/csv",
                Text = ConvertToCsv(data)
            }
        };
    }
}
```

**Benefits for LLMs**:
- **Preview data** before complex operations
- **Validate assumptions** about data structure
- **Show examples** to users when suggesting operations
- **Detect data quality issues** early

**Use Cases**:
- "Show me first 5 rows of Sales data" → Preview before writing formula
- "What columns exist in this table?" → Sample to understand schema
- "Is this query returning data?" → Quick validation

---

### 3. **Enhanced Tool Descriptions** (LOW EFFORT, HIGH VALUE)

**Problem**: Current tool descriptions are brief, lack examples.

**Solution**: Add rich descriptions with examples

**Current**:
```csharp
[McpServerTool(Name = "excel_powerquery", Description = "Manage Power Queries")]
```

**Improved**:
```csharp
[McpServerTool(
    Name = "excel_powerquery",
    Description = @"Manage Power Query M code and data connections.
    
Actions:
- list: Show all queries in workbook
- view: Display M code for a query
- import: Add new query from .pq file
- update: Modify existing query M code
- refresh: Reload query data from source
- set-load-to-table: Configure query to load into worksheet
- delete: Remove query from workbook

Examples:
• List queries: {""action"": ""list"", ""excelPath"": ""data.xlsx""}
• View query: {""action"": ""view"", ""excelPath"": ""data.xlsx"", ""queryName"": ""Sales""}
• Refresh: {""action"": ""refresh"", ""excelPath"": ""data.xlsx"", ""queryName"": ""Sales""}

Note: Use batch sessions for multiple operations on same file."
)]
```

**Benefits**:
- **Self-documenting** - LLMs understand capabilities without external docs
- **Better suggestions** - Examples show LLMs correct usage patterns
- **Reduced errors** - Clear parameter requirements
- **Faster learning** - LLMs adapt to tool usage faster

---

### 4. **Structured Error Responses** (MEDIUM PRIORITY)

**Problem**: Errors return plain strings, LLMs can't programmatically handle them.

**Solution**: Return structured error objects with error codes

**Current**:
```csharp
throw new McpException($"Query '{queryName}' not found");
```

**Improved**:
```csharp
throw new McpException(
    message: $"Query '{queryName}' not found",
    code: "QUERY_NOT_FOUND",
    data: new 
    {
        queryName,
        availableQueries = await ListQueries(excelPath),
        suggestedAction = new
        {
            tool = "excel_powerquery",
            action = "list",
            reason = "See available queries"
        }
    }
);
```

**Benefits for LLMs**:
- **Programmatic handling** - Can detect specific error types
- **Auto-recovery** - Suggested actions help LLMs fix issues
- **Context preservation** - Data field provides debugging info
- **Better UX** - LLMs can explain errors to users more clearly

**Error Code Categories**:
```csharp
public static class ErrorCodes
{
    // Not Found Errors (404-like)
    public const string QUERY_NOT_FOUND = "QUERY_NOT_FOUND";
    public const string WORKSHEET_NOT_FOUND = "WORKSHEET_NOT_FOUND";
    public const string PARAMETER_NOT_FOUND = "PARAMETER_NOT_FOUND";
    
    // Validation Errors (400-like)
    public const string INVALID_QUERY_NAME = "INVALID_QUERY_NAME";
    public const string INVALID_FILE_EXTENSION = "INVALID_FILE_EXTENSION";
    public const string MISSING_REQUIRED_PARAMETER = "MISSING_REQUIRED_PARAMETER";
    
    // Conflict Errors (409-like)
    public const string QUERY_ALREADY_EXISTS = "QUERY_ALREADY_EXISTS";
    public const string PRIVACY_LEVEL_REQUIRED = "PRIVACY_LEVEL_REQUIRED";
    
    // Security Errors (403-like)
    public const string VBA_TRUST_REQUIRED = "VBA_TRUST_REQUIRED";
    
    // System Errors (500-like)
    public const string EXCEL_COM_ERROR = "EXCEL_COM_ERROR";
    public const string FILE_LOCKED = "FILE_LOCKED";
}
```

---

### 5. **~~Progress/Status System~~** (NOT FEASIBLE)

**Problem**: Long-running operations (refresh, import) provide no feedback.

**⚠️ LIMITATION 1**: **Cannot implement** - Excel COM Interop does not provide progress callbacks during operations like `QueryTable.Refresh()`, `Connection.Refresh()`, or workbook saves. These operations are synchronous blocking calls with no intermediate progress reporting.

**⚠️ LIMITATION 2**: **Performance metrics redundant** - As the LLM calling the MCP server, I already measure how long each tool call takes via the MCP protocol itself. Adding duration metrics to responses would duplicate information I already have.

**Alternative Approach**: Time estimates in tool descriptions
```csharp
[McpServerTool(
    Name = "excel_powerquery",
    Description = @"Manage Power Query M code and data connections.
    
Actions: list, view, import, update, refresh, delete

⏱️ Performance Notes:
• Refresh operations: 5-30 seconds (typical), 1-2 minutes (large datasets >100K rows)
• Import/export: <1 second
• Update M code: 1-2 seconds"
)]
public static string ExcelPowerQuery(string action, ...)
```

**What we SHOULD provide**:
- ✅ **Estimated operation times** - Document in tool descriptions to set user expectations
- ✅ **Start/completion messages** - Simple status updates

**What we should NOT provide** (redundant/impossible):
- ❌ **Real-time progress** - Not available from Excel COM API
- ❌ **Performance metrics in responses** - LLMs already measure via MCP protocol

**Recommendation**: Document typical operation durations in tool descriptions to help LLMs set accurate user expectations

---

### 6. **Batch Operation Results** (LOW EFFORT, HIGH VALUE)

**Problem**: Batch operations return individual results, hard to summarize.

**Solution**: Add batch summary to commit response

**Current**:
```json
{
  "success": true,
  "message": "Batch committed successfully"
}
```

**Improved**:
```json
{
  "success": true,
  "message": "Batch committed successfully",
  "summary": {
    "totalOperations": 5,
    "successCount": 5,
    "failureCount": 0,
    "operations": [
      {"action": "pq-import", "queryName": "Sales", "result": "success"},
      {"action": "pq-refresh", "queryName": "Sales", "result": "success"},
      {"action": "sheet-create", "sheetName": "Dashboard", "result": "success"},
      {"action": "range-set-values", "range": "A1:D10", "result": "success"},
      {"action": "file-save", "result": "success"}
    ]
  }
}
```

**Benefits**:
- **LLM summary** - Can report "5/5 operations succeeded"
- **Debugging** - See which operation failed in batch
- **User reporting** - Clear communication of what happened

**Note**: Performance metrics removed (duration, avgOperationTime) - LLMs already track these via MCP protocol response times. No need for redundancy.

---

### 7. **Capability Discovery** (LOW PRIORITY)

**Problem**: LLMs must know tool capabilities upfront.

**Solution**: Add capability introspection

**Implementation**:
```csharp
[McpServerTool(Name = "excel_capabilities")]
public static string GetCapabilities()
{
    return JsonSerializer.Serialize(new
    {
        version = "1.0.0",
        features = new
        {
            batchSessions = true,
            instancePooling = true,
            powerQuery = true,
            vba = true,
            dataModel = true,
            tables = true,
            ranges = true
        },
        limits = new
        {
            maxBatchOperations = 100,
            maxPoolSize = 10,
            maxQueryNameLength = 255
        },
        supportedFileFormats = new[] { ".xlsx", ".xlsm" },
        mcpProtocolVersion = "2024-11-05"
    });
}
```

**Benefits**:
- **Version-aware** - LLMs know what features exist
- **Limit-aware** - LLMs respect constraints
- **Future-proof** - Easy to add new capabilities

---

## Implementation Priority

### Phase 1 (Quick Wins - 1-2 days)
1. ✅ Enhanced tool descriptions (already have good foundation)
2. Structured error codes (ErrorCodes class + McpException updates)
3. Batch operation summaries (update BatchSessionTool.Commit)

### Phase 2 (Medium Effort - 3-5 days)
4. Resource system (workbook introspection)
5. Sampling system (data preview)
6. Capability discovery tool

### Phase 3 (Future - 5-7 days)
7. ~~Progress notifications~~ (NOT FEASIBLE - COM Interop limitation)
8. Advanced resource templates (query schemas, table metadata)
9. Performance metrics in responses (operation duration, etc.)

---

## Example: Enhanced Error Handling Implementation

**Before**:
```csharp
if (query == null)
{
    throw new McpException($"Query '{queryName}' not found");
}
```

**After**:
```csharp
if (query == null)
{
    var availableQueries = await ListPowerQueries(batch);
    
    throw new StructuredMcpException(
        message: $"Query '{queryName}' not found in workbook",
        errorCode: ErrorCodes.QUERY_NOT_FOUND,
        data: new
        {
            queryName,
            excelPath = batch.FilePath,
            availableQueries = availableQueries.Queries.Select(q => q.Name).ToArray(),
            suggestedAction = new
            {
                tool = "excel_powerquery",
                action = "list",
                params = new { excelPath = batch.FilePath },
                rationale = "List all available queries to find the correct name"
            }
        }
    );
}
```

**LLM Workflow**:
```
LLM: excel_powerquery view "Saels" (typo)
MCP: Error QUERY_NOT_FOUND
     - Available: ["Sales", "Customers", "Products"]
     - Suggested: list queries to see all names
LLM: Ah, user meant "Sales" not "Saels"
     excel_powerquery view "Sales"
MCP: Success!
```

---

## Metrics for Success

### For LLMs:
- **Reduced errors** - Fewer invalid tool calls
- **Faster workflows** - Fewer round-trips needed
- **Better suggestions** - More context-aware recommendations
- **Self-service** - Can introspect without docs

### For Users:
- **Clearer feedback** - Better error messages
- **Faster operations** - LLMs make correct calls first try
- **Better UX** - LLMs can explain what they're doing
- **More reliable** - Fewer failed operations

---

## Conclusion

These improvements leverage MCP SDK capabilities to make ExcelMcp more LLM-friendly:

1. **Resources** - Let LLMs discover what exists
2. **Sampling** - Let LLMs preview data
3. **Rich descriptions** - Self-documenting tools
4. **Structured errors** - Programmatic error handling
5. ~~**Progress**~~ - NOT FEASIBLE (COM limitation + LLMs already track performance via MCP protocol)
6. **Batch summaries** - Clear reporting
7. **Capabilities** - Version-aware introspection

**Implementation cost**: ~5-7 days total (reduced from initial 7-10 days estimate)  
**Value to LLMs**: Significantly improved workflow efficiency  
**Backward compatibility**: 100% (all additions, no breaking changes)

**Key Insights**:
- **COM Interop Constraint**: Excel COM API does not provide progress callbacks - operations are synchronous blocking calls
- **MCP Protocol Advantage**: LLMs already measure tool call durations via MCP protocol - no need to duplicate this in responses
- **Focus**: Time estimates in tool descriptions help LLMs set user expectations without redundancy

**Recommendation**: Start with Phase 1 (quick wins) to validate approach, then proceed with Phase 2.

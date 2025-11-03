# Fix MCP Exception Handling - Analysis and Plan

## Problem

Many MCP Server tool methods are returning JSON with `success: false` instead of throwing `McpException`.

**Why this is wrong:**
1. MCP protocol expects exceptions for errors
2. LLMs see HTTP 200 + JSON `{success: false}` which is confusing
3. Error handling in MCP clients doesn't trigger properly
4. Inconsistent with other tools that correctly throw McpException

## Current Pattern (WRONG)

```csharp
private static async Task<string> SomeAction(...)
{
    var result = await ExcelToolsBase.WithBatchAsync(...);
    
    // ❌ NO error check - just returns JSON even if failed!
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

## Correct Pattern

```csharp
private static async Task<string> SomeAction(...)
{
    var result = await ExcelToolsBase.WithBatchAsync(...);
    
    // ✅ Check for errors and throw McpException
    if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
    {
        throw new ModelContextProtocol.McpException($"action-name failed for '{param}': {result.ErrorMessage}");
    }
    
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

## Affected Files

Based on scan:
- ExcelConnectionTool.cs - 8 methods
- ExcelDataModelTool.cs - 14 methods
- ExcelNamedRangeTool.cs - 2+ methods
- ExcelPowerQueryTool.cs - Many methods (some already correct)
- ExcelRangeTool.cs - Many methods
- ExcelVbaTool.cs - Multiple methods
- PivotTableTool.cs - Multiple methods
- TableTool.cs - 21 methods
- ExcelWorksheetTool.cs - Multiple methods

## Fix Strategy

1. **Identify the pattern**: Find all methods that call `WithBatchAsync` or execute Core commands
2. **Add error check**: Insert `if (!result.Success ...)` before `return JsonSerializer.Serialize`
3. **Throw McpException**: With descriptive message including action name and context
4. **Preserve success path**: Only serialize and return JSON when `result.Success == true`

## Implementation

Use PowerShell to:
1. Find all methods returning `JsonSerializer.Serialize(result`
2. Check if they have error handling before return
3. If not, insert the error check pattern

## Example Transformations

### Before
```csharp
private static async Task<string> DeleteTable(TableCommands commands, string filePath, string? tableName, string? batchId)
{
    if (string.IsNullOrWhiteSpace(tableName)) 
        ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

    var result = await ExcelToolsBase.WithBatchAsync(
        batchId,
        filePath,
        save: true,
        async (batch) => await commands.DeleteAsync(batch, tableName!)
    );
    
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions); // ❌ No error check!
}
```

### After
```csharp
private static async Task<string> DeleteTable(TableCommands commands, string filePath, string? tableName, string? batchId)
{
    if (string.IsNullOrWhiteSpace(tableName)) 
        ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

    var result = await ExcelToolsBase.WithBatchAsync(
        batchId,
        filePath,
        save: true,
        async (batch) => await commands.DeleteAsync(batch, tableName!)
    );
    
    // ✅ Check for errors
    if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
    {
        throw new ModelContextProtocol.McpException($"delete failed for table '{tableName}': {result.ErrorMessage}");
    }
    
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

## Testing

After fix:
1. Build should succeed
2. Error cases should throw McpException (HTTP 500) not return JSON (HTTP 200)
3. LLMs should see proper exception messages, not `{success: false, errorMessage: "..."}`

## Benefits for LLMs

✅ **Clear error signals**: HTTP 500 instead of HTTP 200 with error JSON  
✅ **Consistent behavior**: All tools throw exceptions on error  
✅ **Better MCP integration**: Follows MCP protocol expectations  
✅ **Easier debugging**: Exception stack traces instead of buried error messages  
✅ **Proper error handling**: MCP clients can catch and display exceptions correctly

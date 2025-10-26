# MCP Server Tool Conversion Guide

## Overview
Convert all MCP Server tools from synchronous Excel operations to async batch API with optional `batchId` parameter for LLM-controlled batch sessions.

## Conversion Pattern

### 1. Update Tool Method Signature

**Before:**
```csharp
[McpServerTool(Name = "excel_cell")]
[Description("Manage Excel cells...")]
public static string ExcelCell(
    string action,
    string excelPath,
    string sheetName,
    string cellAddress)
```

**After:**
```csharp
[McpServerTool(Name = "excel_cell")]
[Description("Manage Excel cells... Optional batchId for batch sessions.")]
public static async Task<string> ExcelCell(  // ← Add async Task<string>
    string action,
    string excelPath,
    string sheetName,
    string cellAddress,
    [Description("Optional batch session ID from begin_excel_batch")]
    string? batchId = null)  // ← Add optional batchId parameter
```

### 2. Update Private Helper Methods

**Before:**
```csharp
private static string GetValue(CellCommands commands, string filePath, ...)
{
    var result = commands.GetValue(filePath, sheetName, cellAddress);
    return JsonSerializer.Serialize(...);
}
```

**After:**
```csharp
private static async Task<string> GetValueAsync(  // ← async Task<string>
    CellCommands commands, 
    string filePath, 
    ..., 
    string? batchId)  // ← Add batchId parameter
{
    var result = await ExcelToolsBase.WithBatchAsync(
        batchId,
        filePath,
        save: false,  // ← false for read operations
        async (batch) => await commands.GetValueAsync(batch, sheetName, cellAddress));
    
    return JsonSerializer.Serialize(...);
}
```

### 3. Update Method Name Mapping

Core commands now have `Async` suffix:

| Old Name (Sync) | New Name (Async) |
|-----------------|------------------|
| `List()` | `ListAsync()` |
| `GetValue()` | `GetValueAsync()` |
| `SetValue()` | `SetValueAsync()` |
| `Read()` | `ReadAsync()` |
| `Write()` | `WriteAsync()` |
| `Create()` | `CreateAsync()` |
| `Import()` | `ImportAsync()` |
| `Export()` | `ExportAsync()` |
| `Update()` | `UpdateAsync()` |
| `Delete()` | `DeleteAsync()` |
| `Refresh()` | `RefreshAsync()` |
| etc. | etc. |

### 4. Save vs. No-Save Operations

**Read operations (save: false):**
- List
- View/Get
- Read
- Peek

**Write operations (save: true):**
- Create
- Update
- Delete
- Write
- Set
- Import
- Refresh

### 5. Call Chain Update

**Before:**
```csharp
case "get-value":
    return GetValue(cellCommands, excelPath, sheetName, cellAddress);
```

**After:**
```csharp
case "get-value":
    return await GetValueAsync(cellCommands, excelPath, sheetName, cellAddress, batchId);
    // ↑ Add await and batchId parameter
```

## Files to Convert

### High Priority (Most Used)
1. ✅ ExcelFileTool.cs - DONE
2. ❌ ExcelWorksheetTool.cs - ~9 actions
3. ❌ ExcelPowerQueryTool.cs - ~18 actions
4. ❌ ExcelParameterTool.cs - ~6 actions
5. ❌ ExcelCellTool.cs - ~4 actions
6. ❌ ExcelVbaTool.cs - ~7 actions

### Medium Priority
7. ❌ ExcelConnectionTool.cs - ~11 actions
8. ❌ ExcelDataModelTool.cs - ~8 actions

### Low Priority (Less Common)
9. ❌ ExcelVersionTool.cs - Version checking (may not need batch)

## Example: Complete Conversion

### Before (ExcelCellTool.cs)
```csharp
[McpServerTool(Name = "excel_cell")]
public static string ExcelCell(
    string action,
    string excelPath,
    string sheetName,
    string cellAddress,
    string? value = null,
    string? formula = null)
{
    var cellCommands = new CellCommands();
    
    switch (action.ToLowerInvariant())
    {
        case "get-value":
            return GetValue(cellCommands, excelPath, sheetName, cellAddress);
        case "set-value":
            return SetValue(cellCommands, excelPath, sheetName, cellAddress, value!);
        // etc.
    }
}

private static string GetValue(CellCommands commands, string filePath, string sheetName, string cellAddress)
{
    var result = commands.GetValue(filePath, sheetName, cellAddress);
    if (!result.Success)
    {
        throw new ModelContextProtocol.McpException($"get-value failed: {result.ErrorMessage}");
    }
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

### After (ExcelCellTool.cs)
```csharp
[McpServerTool(Name = "excel_cell")]
[Description("... Optional batchId for batch sessions.")]
public static async Task<string> ExcelCell(
    string action,
    string excelPath,
    string sheetName,
    string cellAddress,
    string? value = null,
    string? formula = null,
    [Description("Optional batch session ID from begin_excel_batch")]
    string? batchId = null)
{
    var cellCommands = new CellCommands();
    
    switch (action.ToLowerInvariant())
    {
        case "get-value":
            return await GetValueAsync(cellCommands, excelPath, sheetName, cellAddress, batchId);
        case "set-value":
            return await SetValueAsync(cellCommands, excelPath, sheetName, cellAddress, value!, batchId);
        // etc.
    }
}

private static async Task<string> GetValueAsync(
    CellCommands commands, 
    string filePath, 
    string sheetName, 
    string cellAddress,
    string? batchId)
{
    var result = await ExcelToolsBase.WithBatchAsync(
        batchId,
        filePath,
        save: false,  // Read operation - don't save
        async (batch) => await commands.GetValueAsync(batch, sheetName, cellAddress));
    
    if (!result.Success)
    {
        throw new ModelContextProtocol.McpException($"get-value failed: {result.ErrorMessage}");
    }
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

## Special Cases

### File Creation (ExcelFileTool)
File creation doesn't use batch because it creates a new file:
```csharp
// Don't use WithBatchAsync for file creation
var result = await fileCommands.CreateEmptyAsync(excelPath, overwriteIfExists: false);
```

### Connection Tools
Some connection tools use custom code, not Command classes. May need special handling.

## Testing After Conversion

```bash
# Build MCP Server
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# Should show 0 errors, 0 warnings
```

## Verification Checklist

For each converted tool:
- [ ] Method signature has `async Task<string>` return type
- [ ] Optional `batchId` parameter added
- [ ] Description mentions "Optional batchId for batch sessions"
- [ ] All helper methods have `async Task<string>` return type
- [ ] All Core method calls use `Async` suffix (e.g., `GetValueAsync`)
- [ ] All helper methods use `WithBatchAsync` with correct save flag
- [ ] All `case` statements have `await` before helper call
- [ ] BatchId passed to all helper methods

## Bulk Conversion Script Pattern

For each tool file:
1. Add `batchId` parameter to main tool method
2. Make main tool method `async Task<string>`
3. Add `await` to all case statement helper calls
4. Pass `batchId` to all helper calls
5. Make all helper methods `async Task<string>`
6. Add `batchId` parameter to all helper methods
7. Replace synchronous Core calls with `WithBatchAsync`
8. Update Core method names to `Async` versions

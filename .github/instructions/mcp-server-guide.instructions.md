---
applyTo: "src/ExcelMcp.McpServer/**/*.cs"
---

# MCP Server Development Guide

> **Model Context Protocol server for AI-assisted Excel development**

## Purpose
AI-assisted Excel development (Power Query refactoring, VBA enhancement, code review) - NOT for ETL pipelines.

## Resource-Based Architecture (11 Tools)

| Tool | Purpose |
|------|---------|
| **excel_file** | Excel file operations (create-empty, close-workbook, test) |
| **excel_powerquery** | Power Query lifecycle + optional `privacyLevel` |
| **excel_worksheet** | Worksheet operations (lifecycle only - data via excel_range) |
| **excel_range** | Range operations (values, formulas, clear, copy, hyperlinks) |
| **excel_parameter** | Named ranges |
| **excel_table** | Excel Table (ListObject) operations |
| **excel_connection** | Connection management (OLEDB, ODBC, Text, Web, etc.) |
| **excel_datamodel** | Data Model / Power Pivot operations |
| **excel_vba** | VBA lifecycle + `VbaTrustRequiredResult` |
| **begin_excel_batch** | Start batch session (multi-operation performance) |
| **commit_excel_batch** | Save/discard batch session |

## Implementation Patterns

### Action-Based Routing
```csharp
[McpServerTool]
public async Task<string> ExcelPowerQuery(string action, string excelPath, ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ListPowerQueries(...),
        "view" => ViewPowerQuery(...),
        _ => ThrowUnknownAction(action, "list", "view", ...)
    };
}
```

### Error Handling
```csharp
try
{
    var result = await _commands.Import(excelPath, queryName, mCodeFile);
    if (!result.Success) throw new McpException($"import failed: {result.ErrorMessage}");
    return JsonSerializer.Serialize(result, JsonOptions);
}
catch (McpException) { throw; }
catch (Exception ex) { ThrowInternalError(ex, action, excelPath); throw; }
```

### Async Handling
MCP tools are synchronous, wrap async Core methods:
```csharp
var result = commands.Import(excelPath, queryName, mCodeFile).GetAwaiter().GetResult();
```

### JSON Serialization
```csharp
// ✅ ALWAYS use JsonSerializer
return JsonSerializer.Serialize(result, JsonOptions);

// ❌ NEVER manual JSON strings (path escaping issues)
```

## JSON Deserialization & COM Marshalling

**⚠️ CRITICAL:** MCP deserializes JSON arrays to `JsonElement`, NOT primitives. Excel COM requires proper types.

**Problem:** `values: [["text", 123, true]]` → `List<List<object?>>` where each object is `JsonElement`.

**Solution:** Convert before COM assignment:
```csharp
private static object ConvertToCellValue(object? value)
{
    if (value is System.Text.Json.JsonElement jsonElement)
    {
        return jsonElement.ValueKind switch
        {
            JsonValueKind.String => jsonElement.GetString() ?? string.Empty,
            JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => string.Empty
        };
    }
    return value;
}
```

**When needed:** 2D arrays, nested JSON → COM APIs. **Not needed:** Simple string/int/bool parameters.

## Best Practices

1. **Throw McpException** - Don't return JSON errors
2. **Use async wrappers** - `.GetAwaiter().GetResult()`
3. **Validate parameters** - Use helper methods
4. **Security defaults** - Never auto-apply privacy/trust settings
5. **Update server.json** - Keep synchronized with tool changes
6. **JSON serialization** - Always use `JsonSerializer`
7. **Handle JsonElement** - Convert before COM marshalling

---
applyTo: "src/ExcelMcp.McpServer/**/*.cs"
---

# MCP Server Development Guide

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

## LLM Guidance Development

**See: [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)** for complete guidance on creating guidance for LLMs consuming the MCP server.


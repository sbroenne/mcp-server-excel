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

### Error Handling (MANDATORY PATTERN)

**⚠️ CRITICAL: MCP tools must return JSON responses with `isError: true` for business errors, NOT throw exceptions!**

This follows the official MCP specification which defines two error mechanisms:
1. **Protocol Errors** (JSON-RPC): Unknown tools, invalid arguments → throw exceptions → HTTP error codes
2. **Tool Execution Errors**: Business logic failures → return JSON with `isError: true` → HTTP 200

```csharp
private static async Task<string> SomeActionAsync(Commands commands, string excelPath, string? param, string? batchId)
{
    // 1. Validate parameters (throw McpException for invalid input - PROTOCOL ERROR)
    if (string.IsNullOrEmpty(param))
        throw new ModelContextProtocol.McpException("param is required for action");

    // 2. Call Core Command via WithBatchAsync
    var result = await ExcelToolsBase.WithBatchAsync(
        batchId,
        excelPath,
        save: true,
        async (batch) => await commands.SomeAsync(batch, param));

    // 3. ✅ CORRECT: Always return JSON - let result.Success indicate business errors
    // MCP clients receive: { "success": false, "errorMessage": "...", "isError": true }
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

**When to Throw McpException:**
- ✅ **Parameter validation** - missing required params, invalid formats (pre-conditions)
- ✅ **File not found** - workbook doesn't exist (pre-conditions)
- ✅ **Batch not found** - invalid batch session (pre-conditions)
- ❌ **NOT for business logic errors** - table not found, query failed, connection error, etc.

**Why This Pattern:**
- ✅ MCP spec requires business errors return JSON with `isError: true` flag
- ✅ HTTP 200 + JSON error = client can parse and handle gracefully
- ✅ Core Commands return result objects with `Success` flag - serialize them directly!
- ❌ Throwing exceptions for business errors = harder for MCP clients to handle programmatically

**Example - Business Error (return JSON):**
```csharp
// Core returns: { Success = false, ErrorMessage = "Table 'Sales' not found" }
// MCP Tool: Return this as-is
return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
// Client receives via MCP protocol:
// {
//   "jsonrpc": "2.0",
//   "id": 4,
//   "result": {
//     "content": [{"type": "text", "text": "{\"success\": false, \"errorMessage\": \"Table 'Sales' not found\"}"}],
//     "isError": true
//   }
// }
```

**Example - Validation Error (throw exception):**
```csharp
// Missing required parameter - PROTOCOL ERROR
if (string.IsNullOrWhiteSpace(tableName))
{
    throw new ModelContextProtocol.McpException("tableName is required for create-from-table action");
}
// Client receives: JSON-RPC error with HTTP error code
```

**Reference:** See `critical-rules.instructions.md` Rule 17 for complete guidance and historical context.

**Top-Level Error Handling:**
```csharp
[McpServerTool]
public static async Task<string> ExcelTool(ToolAction action, ...)
{
    try
    {
        return action switch
        {
            ToolAction.Action1 => await Action1Async(...),
            _ => throw new ModelContextProtocol.McpException($"Unknown action: {action}")
        };
    }
    catch (ModelContextProtocol.McpException)
    {
        throw; // Re-throw MCP exceptions as-is
    }
    catch (TimeoutException ex)
    {
        // Enrich timeout errors with operation-specific guidance
        var result = new OperationResult
        {
            Success = false,
            ErrorMessage = ex.Message,
            FilePath = excelPath,
            Action = action.ToActionString(),

            SuggestedNextActions = new List<string>
            {
                "Check if Excel is showing a dialog or prompt",
                "Verify data source connectivity",
                "For large datasets, operation may need more time"
            },

            OperationContext = new Dictionary<string, object>
            {
                { "OperationType", "ToolName.ActionName" },
                { "TimeoutReached", true }
            },

            IsRetryable = !ex.Message.Contains("maximum timeout"),
            RetryGuidance = ex.Message.Contains("maximum timeout")
                ? "Maximum timeout reached. Check connectivity manually."
                : "Retry acceptable if issue is transient."
        };

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
    catch (Exception ex)
    {
        ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
        throw; // Unreachable but satisfies compiler
    }
}
```

### Timeout Exception Enrichment

**When to enrich TimeoutException:**
- Heavy operations: refresh, data model operations, large range operations
- Catch TimeoutException separately from general exceptions
- Return enriched OperationResult with LLM guidance fields

**Pattern:**
```csharp
catch (TimeoutException ex)
{
    var result = new OperationResult
    {
        Success = false,
        ErrorMessage = ex.Message,

        // LLM guidance fields
        SuggestedNextActions = new List<string> { /* operation-specific */ },
        OperationContext = new Dictionary<string, object> { /* diagnostics */ },
        IsRetryable = !ex.Message.Contains("maximum timeout"),
        RetryGuidance = /* retry strategy */
    };

    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

**See:** `docs/TIMEOUT-IMPLEMENTATION-GUIDE.md` for complete examples.

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

1. **✅ ALWAYS return JSON** - Serialize Core Command results directly, let `success` flag indicate errors
2. **Throw McpException sparingly** - Only for parameter validation and pre-conditions, NOT business errors
3. **Validate parameters early** - Throw McpException for missing/invalid params before calling Core Commands
4. **Use async wrappers** - `.GetAwaiter().GetResult()` (deprecated pattern, prefer async)
5. **Security defaults** - Never auto-apply privacy/trust settings
6. **Update server.json** - Keep synchronized with tool changes
7. **JSON serialization** - Always use `JsonSerializer`
8. **Handle JsonElement** - Convert before COM marshalling

## Common Mistakes to Avoid

### ❌ MISTAKE: Throwing Exceptions for Business Errors
```csharp
// ❌ WRONG: Throws exception for business logic errors (violates MCP spec)
var result = await commands.SomeAsync(batch, param);
if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
{
    throw new ModelContextProtocol.McpException($"action failed: {result.ErrorMessage}");
}
return JsonSerializer.Serialize(result, JsonOptions);
```

### ✅ CORRECT: Always Return JSON
```csharp
// ✅ CORRECT: Return JSON for both success and failure
var result = await commands.SomeAsync(batch, param);
return JsonSerializer.Serialize(result, JsonOptions);
// Client receives: {"success": false, "errorMessage": "..."} with isError: true
```

### ❌ MISTAKE: Not Validating Parameters
```csharp
// ❌ WRONG: Missing parameter validation
var result = await commands.SomeAsync(batch, param);  // param might be null!
return JsonSerializer.Serialize(result, JsonOptions);
```

### ✅ CORRECT: Validate Parameters Early
```csharp
// ✅ CORRECT: Validate before calling Core Commands
if (string.IsNullOrWhiteSpace(param))
{
    throw new ModelContextProtocol.McpException("param is required for this action");
}
var result = await commands.SomeAsync(batch, param);
return JsonSerializer.Serialize(result, JsonOptions);
```

## Verification Checklist

Before committing MCP tool changes:

- [ ] Parameter validation throws McpException (pre-conditions)
- [ ] Business errors return JSON with `success: false` (NOT exceptions)
- [ ] All Core Command results are serialized directly
- [ ] Exception messages include context (action name, parameter values)
- [ ] Build passes with 0 warnings
- [ ] No `if (!result.Success) throw McpException` blocks (violates MCP spec)

## LLM Guidance Development

**See: [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)** for complete guidance on creating guidance for LLMs consuming the MCP server.


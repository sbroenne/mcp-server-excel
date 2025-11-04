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

**⚠️ CRITICAL: Every method that calls Core Commands MUST check result.Success before serializing!**

```csharp
private static async Task<string> SomeActionAsync(Commands commands, string excelPath, string? param, string? batchId)
{
    // 1. Validate parameters (throw McpException for invalid input)
    if (string.IsNullOrEmpty(param))
        throw new ModelContextProtocol.McpException("param is required for action");

    // 2. Call Core Command via WithBatchAsync
    var result = await ExcelToolsBase.WithBatchAsync(
        batchId,
        excelPath,
        save: true,
        async (batch) => await commands.SomeAsync(batch, param));

    // 3. ✅ MANDATORY: Check result.Success BEFORE serializing
    if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
    {
        throw new ModelContextProtocol.McpException($"action failed for '{param}': {result.ErrorMessage}");
    }

    // 4. Only serialize on success
    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
}
```

**Why Critical:**
- ❌ **WRONG**: Returning JSON with `success: false` → HTTP 200 (confuses LLMs)
- ✅ **CORRECT**: Throwing McpException → HTTP 500 (clear error signal)

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

1. **✅ ALWAYS check result.Success** - Before serializing, throw McpException if failed
2. **Throw McpException for errors** - Don't return JSON with success=false
3. **Validate parameters early** - Throw McpException for missing/invalid params
4. **Use async wrappers** - `.GetAwaiter().GetResult()` (deprecated pattern, prefer async)
5. **Security defaults** - Never auto-apply privacy/trust settings
6. **Update server.json** - Keep synchronized with tool changes
7. **JSON serialization** - Always use `JsonSerializer`
8. **Handle JsonElement** - Convert before COM marshalling

## Common Mistakes to Avoid

### ❌ MISTAKE: Missing Error Check
```csharp
// ❌ WRONG: Serializes even when result.Success = false
var result = await commands.SomeAsync(batch, param);
return JsonSerializer.Serialize(result, JsonOptions); // HTTP 200 with error JSON!
```

### ✅ CORRECT: Always Check Before Serializing
```csharp
// ✅ CORRECT: Throw exception if operation failed
var result = await commands.SomeAsync(batch, param);
if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
{
    throw new ModelContextProtocol.McpException($"action failed: {result.ErrorMessage}");
}
return JsonSerializer.Serialize(result, JsonOptions); // HTTP 200 only on success!
```

### ❌ MISTAKE: Empty Success Blocks
```csharp
// ❌ WRONG: Useless empty block, no error handling
if (result.Success)
{
    // Empty - does nothing!
}
return JsonSerializer.Serialize(result, JsonOptions); // Still returns on failure!
```

### ✅ CORRECT: Check Failure, Not Success
```csharp
// ✅ CORRECT: Check for failure and throw
if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
{
    throw new ModelContextProtocol.McpException($"action failed: {result.ErrorMessage}");
}
return JsonSerializer.Serialize(result, JsonOptions);
```

## Verification Checklist

Before committing MCP tool changes:

- [ ] Every method calling Core Commands has error check
- [ ] Error check comes BEFORE `JsonSerializer.Serialize`
- [ ] Exception message includes context (action name, parameter values)
- [ ] No empty `if (result.Success) {}` blocks
- [ ] Build passes with 0 warnings
- [ ] Coverage: error checks ≥ serializations (use script below)

**Coverage Check Script:**
```powershell
# Run from repository root
$file = "src/ExcelMcp.McpServer/Tools/YourTool.cs"
$content = Get-Content $file -Raw
$errorChecks = ([regex]::Matches($content, 'if\s*\(\s*!result\.Success')).Count
$serializes = ([regex]::Matches($content, 'JsonSerializer\.Serialize\(result')).Count
Write-Host "Coverage: $errorChecks / $serializes = $([math]::Round(($errorChecks/$serializes)*100,0))%"
# Should be 100% or higher
```

## LLM Guidance Development

**See: [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)** for complete guidance on creating guidance for LLMs consuming the MCP server.


---
applyTo: "src/ExcelMcp.McpServer/**/*.cs"
---

# MCP Server Development Guide

> **All action methods are synchronous.** Tool signatures are `async Task<string>` only to satisfy MCP SDK requirements, but action method implementations use NO async/await.

## Implementation Patterns

### Action-Based Routing

```csharp
[McpServerTool]
public async Task<string> ExcelPowerQuery(string action, string excelPath, ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => List(...),      // Synchronous methods
        "view" => View(...),      // No await!
        _ => ThrowUnknownAction(action, "list", "view", ...)
    };
}
```

### Error Handling (MANDATORY)

**⚠️ CRITICAL: MCP tools must return JSON responses with `isError: true` for business errors, NOT throw exceptions!**

This follows the official MCP specification which defines two error mechanisms:

1. **Protocol Errors** (JSON-RPC): Unknown tools, invalid arguments → throw exceptions → HTTP error codes
2. **Tool Execution Errors**: Business logic failures → return JSON with `isError: true` → HTTP 200

```csharp
private static string SomeAction(Commands commands, string excelPath, string? param, string? batchId)
{
    // 1. Validate parameters (throw McpException for invalid input - PROTOCOL ERROR)
    if (string.IsNullOrEmpty(param))
        throw new ModelContextProtocol.McpException("param is required for action");

    // 2. Call Core Command via WithSession
    var result = ExcelToolsBase.WithSession(
        batchId,
        batch => commands.Some(batch, param));

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
            ToolAction.Action1 => Action1(...),  // Synchronous methods
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

### Method Signatures

**CRITICAL:** All MCP tool action methods are **synchronous** (no async/await). The tool signature is async only to satisfy MCP SDK requirements:

```csharp
// Tool signature: async Task<string> (MCP SDK requirement)
[McpServerTool]
public static async Task<string> ExcelPowerQuery(string action, ...)
{
    // Action methods: synchronous (no await!)
    return action.ToLowerInvariant() switch
    {
        "list" => List(...),        // ✅ Synchronous
        "view" => View(...),        // ✅ Synchronous
        _ => throw new McpException("Unknown action")
    };
}

// Action methods are synchronous:
private static string List(Commands commands, string sessionId)
{
    var result = ExcelToolsBase.WithSession(sessionId,
        batch => commands.List(batch));  // No await needed
    return JsonSerializer.Serialize(result, JsonOptions);
}
```

### JSON Serialization

```csharp
// ✅ ALWAYS use JsonSerializer
return JsonSerializer.Serialize(result, JsonOptions);

// ❌ NEVER manual JSON strings
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

1. **ALWAYS return JSON** - Serialize Core Command results directly, let `success` flag indicate errors
2. **Throw McpException sparingly** - Only for parameter validation and pre-conditions, NOT business errors
3. **Validate parameters early** - Throw McpException for missing/invalid params before calling Core Commands
4. **Action methods are synchronous** - Do NOT use async/await in action method implementations
5. **Security defaults** - Never auto-apply privacy/trust settings
6. **Update server.json** - Keep synchronized with tool changes
7. **JSON serialization** - Always use `JsonSerializer`
8. **Handle JsonElement** - Convert before COM marshalling
9. **Error messages: facts not guidance** - State what failed, not what to do next. LLMs figure out next steps.
10. **NO EMOJIS** - Never use emoji characters in XML comments, `[Description]` attributes, or any code documentation. Use plain text markers like "IMPORTANT:", "WARNING:", "NOTE:" instead.

## Error Message Style

**❌ WRONG: Verbose guidance (LLM doesn't need step-by-step instructions)**
```csharp
errorMessage = "Operation failed. This usually means: (1) Sheet doesn't exist, (2) Range invalid, or (3) Session closed. " +
               "Use excel_worksheet(action: 'list') to verify sheet exists, then excel_file(action: 'list') to check sessions.";
```

**✅ CORRECT: State facts (LLM determines next action)**
```csharp
errorMessage = $"Cannot read range '{range}' on sheet '{sheet}': {ex.Message}";
```

**Why:** LLMs are intelligent agents that determine workflow. Error messages should report what failed and why, not prescribe solutions.

## Common Mistakes to Avoid

### ❌ MISTAKE: Throwing Exceptions for Business Errors
```csharp
// ❌ WRONG: Throws exception for business logic errors (violates MCP spec)
var result = commands.Some(batch, param);
if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
{
    throw new ModelContextProtocol.McpException($"action failed: {result.ErrorMessage}");
}
return JsonSerializer.Serialize(result, JsonOptions);
```

### ✅ CORRECT: Always Return JSON
```csharp
// ✅ CORRECT: Return JSON for both success and failure
var result = commands.Some(batch, param);
return JsonSerializer.Serialize(result, JsonOptions);
// Client receives: {"success": false, "errorMessage": "..."} with isError: true
```

### ❌ MISTAKE: Not Validating Parameters
```csharp
// ❌ WRONG: Missing parameter validation
var result = commands.Some(batch, param);  // param might be null!
return JsonSerializer.Serialize(result, JsonOptions);
```

### ✅ CORRECT: Validate Parameters Early
```csharp
// ✅ CORRECT: Validate before calling Core Commands
if (string.IsNullOrWhiteSpace(param))
{
    throw new ModelContextProtocol.McpException("param is required for this action");
}
var result = commands.Some(batch, param);
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
- [ ] **Tool `[Description]` attribute documents server-specific behavior**
- [ ] **Non-enum parameter values explained (loadDestination, formatCode, etc.)**
- [ ] **Performance guidance (batch mode) is accurate**
- [ ] **Related tools referenced correctly**

## Tool Description vs Prompt Files

**Two types of LLM guidance:**

1. **Tool Descriptions** (`[Description]` attributes in C# code):
   - Part of MCP tool schema sent automatically
   - LLMs see when browsing available tools
   - Brief, action-oriented reference
   - **ALWAYS visible** - shown every time tool is considered
   - Must be kept synchronized with actual tool behavior

2. **Prompt Files** (`.md` files in `Prompts/Content/`):
   - Exposed as separate MCP Prompts via `[McpServerPrompt]`
   - LLMs request explicitly when needed
   - Detailed workflows, checklists, disambiguation
   - **On-demand** - loaded only when LLM requests the prompt

**Critical:** When changing tool behavior, update BOTH:
- The `[Description]` attribute on the tool method
- The corresponding `.md` prompt file (if exists)

### Keeping Descriptions Up-to-Date

**When updating a tool, verify:**
1. ✅ Tool purpose and use cases are clear
2. ✅ Server-specific behavior is documented (defaults, quirks, important notes)
3. ✅ Performance guidance (batch mode) is accurate
4. ✅ Related tools referenced correctly
5. ✅ Non-enum parameter guidance is complete (loadDestination options, format codes, etc.)

**What NOT to include in descriptions:**
- ❌ **Enum action lists** - MCP SDK auto-generates enum values in schema (LLMs see them automatically)
- ❌ **Parameter types** - Schema provides this
- ❌ **Required/optional flags** - Schema provides this

**Example - Good tool description:**
```csharp
[Description(@"Manage Power Query M code and data loading.

LOAD DESTINATIONS (non-enum parameter):
- 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
- 'data-model': Load to Power Pivot Data Model (ready for DAX measures/relationships)
- 'both': Load to BOTH worksheet AND Data Model
- 'connection-only': Don't load data (M code imported but not executed)

TIMEOUT: Long-running refresh/load operations auto-timeout after 5 minutes.

Use excel_datamodel tool for DAX measures after loading to Data Model.")]
```
✅ Describes purpose and use cases
✅ Documents server-specific defaults
✅ Explains non-enum parameter values
✅ References related tools
❌ Does NOT list enum actions (SDK provides)

## LLM Guidance Development

**See: [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)** for complete guidance on creating guidance for LLMs consuming the MCP server.


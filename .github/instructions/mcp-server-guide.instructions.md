---------

applyTo: "src/ExcelMcp.McpServer/**/*.cs"

---applyTo: "src/ExcelMcp.McpServer/**/*.cs"applyTo: "src/ExcelMcp.McpServer/**/*.cs"



# MCP Server Development Guide------

> **Timeout Handling Removed (2025-11-15).** Previous timeout parameters and `TimeoutException` enrichment patterns have been retired. Do **not** add timeout-specific arguments, exception handling, or documentation. Any remaining references in this file are historical and will be cleaned up incrementally.



## Implementation Patterns



### Action-Based Routing# MCP Server Development Guide# MCP Server Development Guide

```csharp

[McpServerTool]

public async Task<string> ExcelPowerQuery(string action, string excelPath, ...)

{## Implementation Patterns## Implementation Patterns

    return action.ToLowerInvariant() switch

    {

        "list" => ListPowerQueries(...),

        "view" => ViewPowerQuery(...),### Action-Based Routing### Action-Based Routing

        _ => ThrowUnknownAction(action, "list", "view", ...)

    };```csharp```csharp

}

```[McpServerTool][McpServerTool]



### Error Handling (MANDATORY)public async Task<string> ExcelPowerQuery(string action, string excelPath, ...)public async Task<string> ExcelPowerQuery(string action, string excelPath, ...)



**⚠️ CRITICAL: Return JSON for business errors, throw McpException only for protocol errors!**{{



```csharp    return action.ToLowerInvariant() switch    return action.ToLowerInvariant() switch

private static async Task<string> SomeActionAsync(Commands commands, string excelPath, string? param, string? batchId)

{    {    {

    // 1. Validate parameters - throw for protocol errors

    if (string.IsNullOrEmpty(param))        "list" => ListPowerQueries(...),        "list" => ListPowerQueries(...),

        throw new ModelContextProtocol.McpException("param is required");

        "view" => ViewPowerQuery(...),        "view" => ViewPowerQuery(...),

    // 2. Call Core Command

    var result = await ExcelToolsBase.WithBatchAsync(        _ => ThrowUnknownAction(action, "list", "view", ...)        _ => ThrowUnknownAction(action, "list", "view", ...)

        batchId, excelPath, save: true,

        async (batch) => await commands.SomeAsync(batch, param));    };    };



    // 3. Return JSON - let result.Success indicate business errors}}

    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);

}``````

```



**When to Throw McpException:**

- ✅ Parameter validation (missing/invalid)### Error Handling (MANDATORY)### Error Handling (MANDATORY PATTERN)

- ✅ File not found

- ✅ Batch not found

- ❌ **NOT** for business errors (table not found, query failed, etc.)

**⚠️ CRITICAL: Return JSON for business errors, throw McpException only for protocol errors!****⚠️ CRITICAL: MCP tools must return JSON responses with `isError: true` for business errors, NOT throw exceptions!**

**Top-Level Error Handling:**

```csharp

[McpServerTool]

public static async Task<string> ExcelTool(ToolAction action, ...)```csharpThis follows the official MCP specification which defines two error mechanisms:

{

    tryprivate static async Task<string> SomeActionAsync(Commands commands, string excelPath, string? param, string? batchId)1. **Protocol Errors** (JSON-RPC): Unknown tools, invalid arguments → throw exceptions → HTTP error codes

    {

        return action switch{2. **Tool Execution Errors**: Business logic failures → return JSON with `isError: true` → HTTP 200

        {

            ToolAction.Action1 => await Action1Async(...),    // 1. Validate parameters - throw for protocol errors

            _ => throw new ModelContextProtocol.McpException($"Unknown action: {action}")

        };    if (string.IsNullOrEmpty(param))```csharp

    }

    catch (ModelContextProtocol.McpException)        throw new ModelContextProtocol.McpException("param is required");private static async Task<string> SomeActionAsync(Commands commands, string excelPath, string? param, string? batchId)

    {

        throw;{

    }

    catch (TimeoutException ex)    // 2. Call Core Command    // 1. Validate parameters (throw McpException for invalid input - PROTOCOL ERROR)

    {

        // Enrich with LLM guidance    var result = await ExcelToolsBase.WithBatchAsync(    if (string.IsNullOrEmpty(param))

        var result = new OperationResult

        {        batchId, excelPath, save: true,        throw new ModelContextProtocol.McpException("param is required for action");

            Success = false,

            ErrorMessage = ex.Message,        async (batch) => await commands.SomeAsync(batch, param));

            SuggestedNextActions = new List<string> { /* operation-specific */ },

            IsRetryable = !ex.Message.Contains("maximum timeout")    // 2. Call Core Command via WithBatchAsync

        };

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);    // 3. Return JSON - let result.Success indicate business errors    var result = await ExcelToolsBase.WithBatchAsync(

    }

    catch (Exception ex)    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);        batchId,

    {

        ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);}        excelPath,

        throw;

    }```        save: true,

}

```        async (batch) => await commands.SomeAsync(batch, param));



### Timeout Exception Enrichment**When to Throw McpException:**



**When:** Heavy operations (refresh, data model, large ranges)- ✅ Parameter validation (missing/invalid)    // 3. ✅ CORRECT: Always return JSON - let result.Success indicate business errors



**Pattern:**- ✅ File not found    // MCP clients receive: { "success": false, "errorMessage": "...", "isError": true }

```csharp

catch (TimeoutException ex)- ✅ Batch not found    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);

{

    var result = new OperationResult- ❌ **NOT** for business errors (table not found, query failed, etc.)}

    {

        Success = false,```

        ErrorMessage = ex.Message,

        SuggestedNextActions = new List<string> { /* diagnostics */ },**Top-Level Error Handling:**

        OperationContext = new Dictionary<string, object> { /* details */ },

        IsRetryable = !ex.Message.Contains("maximum timeout")```csharp**When to Throw McpException:**

    };

    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);[McpServerTool]- ✅ **Parameter validation** - missing required params, invalid formats (pre-conditions)

}

```public static async Task<string> ExcelTool(ToolAction action, ...)- ✅ **File not found** - workbook doesn't exist (pre-conditions)



**See:** `docs/TIMEOUT-IMPLEMENTATION-GUIDE.md`{- ✅ **Batch not found** - invalid batch session (pre-conditions)



### Async Handling    try- ❌ **NOT for business logic errors** - table not found, query failed, connection error, etc.

MCP tools are synchronous, wrap async Core methods:

```csharp    {

var result = commands.Import(excelPath, queryName, mCodeFile).GetAwaiter().GetResult();

```        return action switch**Why This Pattern:**



### JSON Serialization        {- ✅ MCP spec requires business errors return JSON with `isError: true` flag

```csharp

// ✅ ALWAYS use JsonSerializer            ToolAction.Action1 => await Action1Async(...),- ✅ HTTP 200 + JSON error = client can parse and handle gracefully

return JsonSerializer.Serialize(result, JsonOptions);

            _ => throw new ModelContextProtocol.McpException($"Unknown action: {action}")- ✅ Core Commands return result objects with `Success` flag - serialize them directly!

// ❌ NEVER manual JSON strings

```        };- ❌ Throwing exceptions for business errors = harder for MCP clients to handle programmatically



## JSON Deserialization & COM Marshalling    }



**⚠️ CRITICAL:** MCP deserializes JSON arrays to `JsonElement`, NOT primitives. Excel COM requires proper types.    catch (ModelContextProtocol.McpException)**Example - Business Error (return JSON):**



**Problem:** `values: [["text", 123, true]]` → `List<List<object?>>` where each object is `JsonElement`.    {```csharp



**Solution:** Convert before COM assignment:        throw;// Core returns: { Success = false, ErrorMessage = "Table 'Sales' not found" }

```csharp

private static object ConvertToCellValue(object? value)    }// MCP Tool: Return this as-is

{

    if (value is System.Text.Json.JsonElement jsonElement)    catch (TimeoutException ex)return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);

    {

        return jsonElement.ValueKind switch    {// Client receives via MCP protocol:

        {

            JsonValueKind.String => jsonElement.GetString() ?? string.Empty,        // Enrich with LLM guidance// {

            JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),

            JsonValueKind.True => true,        var result = new OperationResult//   "jsonrpc": "2.0",

            JsonValueKind.False => false,

            _ => string.Empty        {//   "id": 4,

        };

    }            Success = false,//   "result": {

    return value;

}            ErrorMessage = ex.Message,//     "content": [{"type": "text", "text": "{\"success\": false, \"errorMessage\": \"Table 'Sales' not found\"}"}],

```

            SuggestedNextActions = new List<string> { /* operation-specific */ },//     "isError": true

**When needed:** 2D arrays, nested JSON → COM APIs

            IsRetryable = !ex.Message.Contains("maximum timeout")//   }

## Best Practices

        };// }

1. **Return JSON** - Serialize Core Command results, let `success` flag indicate errors

2. **Throw McpException sparingly** - Only for parameter validation and pre-conditions        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);```

3. **Validate early** - Check params before calling Core Commands

4. **Security defaults** - Never auto-apply privacy/trust settings    }

5. **Update server.json** - Keep synchronized with tool changes

6. **Handle JsonElement** - Convert before COM marshalling    catch (Exception ex)**Example - Validation Error (throw exception):**



## Tool Description vs Prompt Files    {```csharp



**Two types of LLM guidance:**        ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);// Missing required parameter - PROTOCOL ERROR



1. **Tool Descriptions** (`[Description]` attributes):        throw;if (string.IsNullOrWhiteSpace(tableName))

   - Part of MCP schema sent automatically

   - Brief, action-oriented reference    }{

   - **ALWAYS visible**

}    throw new ModelContextProtocol.McpException("tableName is required for create-from-table action");

2. **Prompt Files** (`.md` in `Prompts/Content/`):

   - Separate MCP Prompts via `[McpServerPrompt]````}

   - Detailed workflows, checklists

   - **On-demand**// Client receives: JSON-RPC error with HTTP error code



**When updating tool, update BOTH**### Timeout Exception Enrichment```



### Keeping Descriptions Up-to-Date



**Verify:****When:** Heavy operations (refresh, data model, large ranges)**Reference:** See `critical-rules.instructions.md` Rule 17 for complete guidance and historical context.

1. Tool purpose clear

2. Server-specific behavior documented (defaults, quirks)

3. Performance guidance accurate

4. Related tools referenced**Pattern:****Top-Level Error Handling:**

5. Non-enum parameter values explained (loadDestination, format codes)

```csharp```csharp

**Don't include:**

- ❌ Enum action lists (SDK auto-generates)catch (TimeoutException ex)[McpServerTool]

- ❌ Parameter types (schema provides)

- ❌ Required/optional flags (schema provides){public static async Task<string> ExcelTool(ToolAction action, ...)



**Example:**    var result = new OperationResult{

```csharp

[Description(@"Manage Power Query M code and data loading.    {    try



⚡ Use begin_excel_batch for 2+ operations (75-90% faster)        Success = false,    {



LOAD DESTINATIONS (non-enum):        ErrorMessage = ex.Message,        return action switch

- 'worksheet': Load to worksheet (DEFAULT)

- 'data-model': Load to Power Pivot Data Model        SuggestedNextActions = new List<string> { /* diagnostics */ },        {

- 'both': Load to BOTH

- 'connection-only': Don't load data        OperationContext = new Dictionary<string, object> { /* details */ },            ToolAction.Action1 => await Action1Async(...),



Use excel_datamodel tool for DAX measures.")]        IsRetryable = !ex.Message.Contains("maximum timeout")            _ => throw new ModelContextProtocol.McpException($"Unknown action: {action}")

```

    };        };

## Verification Checklist

    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);    }

- [ ] Parameter validation throws McpException

- [ ] Business errors return JSON with `success: false`}    catch (ModelContextProtocol.McpException)

- [ ] All Core Command results serialized

- [ ] Exception messages include context```    {

- [ ] Build passes (0 warnings)

- [ ] Tool `[Description]` documents server-specific behavior        throw; // Re-throw MCP exceptions as-is

- [ ] Non-enum parameter values explained

**See:** `docs/TIMEOUT-IMPLEMENTATION-GUIDE.md`    }

**See:** [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)

    catch (TimeoutException ex)

### Async Handling    {

MCP tools are synchronous, wrap async Core methods:        // Enrich timeout errors with operation-specific guidance

```csharp        var result = new OperationResult

var result = commands.Import(excelPath, queryName, mCodeFile).GetAwaiter().GetResult();        {

```            Success = false,

            ErrorMessage = ex.Message,

### JSON Serialization            FilePath = excelPath,

```csharp            Action = action.ToActionString(),

// ✅ ALWAYS use JsonSerializer

return JsonSerializer.Serialize(result, JsonOptions);            SuggestedNextActions = new List<string>

            {

// ❌ NEVER manual JSON strings                "Check if Excel is showing a dialog or prompt",

```                "Verify data source connectivity",

                "For large datasets, operation may need more time"

## JSON Deserialization & COM Marshalling            },



**⚠️ CRITICAL:** MCP deserializes JSON arrays to `JsonElement`, NOT primitives. Excel COM requires proper types.            OperationContext = new Dictionary<string, object>

            {

**Problem:** `values: [["text", 123, true]]` → `List<List<object?>>` where each object is `JsonElement`.                { "OperationType", "ToolName.ActionName" },

                { "TimeoutReached", true }

**Solution:** Convert before COM assignment:            },

```csharp

private static object ConvertToCellValue(object? value)            IsRetryable = !ex.Message.Contains("maximum timeout"),

{            RetryGuidance = ex.Message.Contains("maximum timeout")

    if (value is System.Text.Json.JsonElement jsonElement)                ? "Maximum timeout reached. Check connectivity manually."

    {                : "Retry acceptable if issue is transient."

        return jsonElement.ValueKind switch        };

        {

            JsonValueKind.String => jsonElement.GetString() ?? string.Empty,        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);

            JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),    }

            JsonValueKind.True => true,    catch (Exception ex)

            JsonValueKind.False => false,    {

            _ => string.Empty        ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);

        };        throw; // Unreachable but satisfies compiler

    }    }

    return value;}

}```

```

### Timeout Exception Enrichment

**When needed:** 2D arrays, nested JSON → COM APIs

**When to enrich TimeoutException:**

## Best Practices- Heavy operations: refresh, data model operations, large range operations

- Catch TimeoutException separately from general exceptions

1. **Return JSON** - Serialize Core Command results, let `success` flag indicate errors- Return enriched OperationResult with LLM guidance fields

2. **Throw McpException sparingly** - Only for parameter validation and pre-conditions

3. **Validate early** - Check params before calling Core Commands**Pattern:**

4. **Security defaults** - Never auto-apply privacy/trust settings```csharp

5. **Update server.json** - Keep synchronized with tool changescatch (TimeoutException ex)

6. **Handle JsonElement** - Convert before COM marshalling{

    var result = new OperationResult

## Tool Description vs Prompt Files    {

        Success = false,

**Two types of LLM guidance:**        ErrorMessage = ex.Message,



1. **Tool Descriptions** (`[Description]` attributes):        // LLM guidance fields

   - Part of MCP schema sent automatically        SuggestedNextActions = new List<string> { /* operation-specific */ },

   - Brief, action-oriented reference        OperationContext = new Dictionary<string, object> { /* diagnostics */ },

   - **ALWAYS visible**        IsRetryable = !ex.Message.Contains("maximum timeout"),

        RetryGuidance = /* retry strategy */

2. **Prompt Files** (`.md` in `Prompts/Content/`):    };

   - Separate MCP Prompts via `[McpServerPrompt]`

   - Detailed workflows, checklists    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);

   - **On-demand**}

```

**When updating tool, update BOTH**

**See:** `docs/TIMEOUT-IMPLEMENTATION-GUIDE.md` for complete examples.

### Keeping Descriptions Up-to-Date

### Async Handling

**Verify:**MCP tools are synchronous, wrap async Core methods:

1. Tool purpose clear```csharp

2. Server-specific behavior documented (defaults, quirks)var result = commands.Import(excelPath, queryName, mCodeFile).GetAwaiter().GetResult();

3. Performance guidance accurate```

4. Related tools referenced

5. Non-enum parameter values explained (loadDestination, format codes)### JSON Serialization

```csharp

**Don't include:**// ✅ ALWAYS use JsonSerializer

- ❌ Enum action lists (SDK auto-generates)return JsonSerializer.Serialize(result, JsonOptions);

- ❌ Parameter types (schema provides)

- ❌ Required/optional flags (schema provides)// ❌ NEVER manual JSON strings (path escaping issues)

```

**Example:**

```csharp## JSON Deserialization & COM Marshalling

[Description(@"Manage Power Query M code and data loading.

**⚠️ CRITICAL:** MCP deserializes JSON arrays to `JsonElement`, NOT primitives. Excel COM requires proper types.

⚡ Use begin_excel_batch for 2+ operations (75-90% faster)

**Problem:** `values: [["text", 123, true]]` → `List<List<object?>>` where each object is `JsonElement`.

LOAD DESTINATIONS (non-enum):

- 'worksheet': Load to worksheet (DEFAULT)**Solution:** Convert before COM assignment:

- 'data-model': Load to Power Pivot Data Model```csharp

- 'both': Load to BOTHprivate static object ConvertToCellValue(object? value)

- 'connection-only': Don't load data{

    if (value is System.Text.Json.JsonElement jsonElement)

Use excel_datamodel tool for DAX measures.")]    {

```        return jsonElement.ValueKind switch

        {

## Verification Checklist            JsonValueKind.String => jsonElement.GetString() ?? string.Empty,

            JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),

- [ ] Parameter validation throws McpException            JsonValueKind.True => true,

- [ ] Business errors return JSON with `success: false`            JsonValueKind.False => false,

- [ ] All Core Command results serialized            _ => string.Empty

- [ ] Exception messages include context        };

- [ ] Build passes (0 warnings)    }

- [ ] Tool `[Description]` documents server-specific behavior    return value;

- [ ] Non-enum parameter values explained}

```

**See:** [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)

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
9. **Error messages: facts not guidance** - State what failed, not what to do next. LLMs figure out next steps.

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

⚡ PERFORMANCE: Use begin_excel_batch for 2+ operations (75-90% faster)

LOAD DESTINATIONS (non-enum parameter):
- 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
- 'data-model': Load to Power Pivot Data Model (ready for DAX measures/relationships)
- 'both': Load to BOTH worksheet AND Data Model
- 'connection-only': Don't load data (M code imported but not executed)

Use excel_datamodel tool for DAX measures after loading to Data Model.")]
```
✅ Describes purpose and use cases
✅ Documents server-specific defaults
✅ Explains non-enum parameter values
✅ References related tools
❌ Does NOT list enum actions (SDK provides)

## LLM Guidance Development

**See: [mcp-llm-guidance.instructions.md](mcp-llm-guidance.instructions.md)** for complete guidance on creating guidance for LLMs consuming the MCP server.


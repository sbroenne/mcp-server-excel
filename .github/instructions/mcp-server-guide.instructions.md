---
applyTo: "src/ExcelMcp.McpServer/**/*.cs"
---

# MCP Server Development Guide

> **Model Context Protocol server for AI-assisted Excel development**

## What is the MCP Server?

ExcelMcp includes an **MCP server** that transforms CLI commands into conversational development workflows for AI assistants like GitHub Copilot.

**Purpose:** AI-assisted Excel development (Power Query refactoring, VBA enhancement, code review)  
**NOT for:** ETL pipelines or data processing

---

## Starting the Server

```powershell
dotnet run --project src/ExcelMcp.McpServer
```

---

## Resource-Based Architecture (6 Tools)

### 1. excel_file
- **Actions:** create-empty
- **Purpose:** Excel-specific file creation

### 2. excel_powerquery
- **Actions:** list, view, import, export, update, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config
- **Purpose:** Complete Power Query lifecycle
- **Security:** Optional `privacyLevel` parameter (None, Private, Organizational, Public)

### 3. excel_worksheet
- **Actions:** list, read, write, create, rename, copy, delete, clear, append
- **Purpose:** Full worksheet lifecycle with bulk data operations

### 4. excel_parameter
- **Actions:** list, get, set, create, delete
- **Purpose:** Named ranges as configuration parameters

### 5. excel_cell
- **Actions:** get-value, set-value, get-formula, set-formula
- **Purpose:** Granular cell control

### 6. excel_vba
- **Actions:** list, export, import, update, run, delete
- **Purpose:** Complete VBA lifecycle
- **Security:** Returns `VbaTrustRequiredResult` when VBA trust not enabled

---

## Development Use Cases

### Power Query Development
```text
Developer: "This Power Query is slow. Can you refactor it?"
Copilot: [excel_powerquery view → analyzes M code → excel_powerquery update with optimized code]
```

### Power Query with Privacy (Security-First)
```text
Developer: "Import this query that combines data sources"
Copilot: [Attempts excel_powerquery import without privacyLevel]
         [Receives PowerQueryPrivacyErrorResult]
         "This combines data sources. Excel requires a privacy level. I recommend 'Private'. Proceed?"
Developer: "Yes"
Copilot: [excel_powerquery import with privacyLevel="Private"]
```

### VBA Enhancement
```text
Developer: "Add error handling to this VBA module"
Copilot: [excel_vba export → enhances with try-catch → excel_vba update]
```

### VBA with Trust Guidance
```text
Developer: "List VBA modules"
Copilot: [Attempts excel_vba list]
         [Receives VbaTrustRequiredResult]
         "VBA trust not enabled. Configure manually in Excel:
          1. File → Options → Trust Center
          2. Macro Settings
          3. Check 'Trust access to VBA project object model'"
```

---

## Tool Implementation Pattern

### Action-Based Routing
```csharp
[McpServerTool]
public async Task<string> ExcelPowerQuery(
    string action,
    string excelPath,
    string? queryName = null,
    string? mCodeFile = null,
    string? privacyLevel = null)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ListPowerQueries(powerQueryCommands, excelPath),
        "view" => ViewPowerQuery(powerQueryCommands, excelPath, queryName),
        "import" => await ImportPowerQuery(powerQueryCommands, excelPath, queryName, mCodeFile, privacyLevel),
        "update" => await UpdatePowerQuery(powerQueryCommands, excelPath, queryName, mCodeFile, privacyLevel),
        _ => ThrowUnknownAction(action, "list", "view", "import", "update", "delete", ...)
    };
}
```

### Error Handling (MCP SDK Pattern)
```csharp
try
{
    // Call Core business logic
    var result = await _commands.Import(excelPath, queryName, mCodeFile);
    
    // Check result
    if (!result.Success)
    {
        throw new McpException($"import failed for '{excelPath}': {result.ErrorMessage}");
    }
    
    return JsonSerializer.Serialize(result, JsonOptions);
}
catch (McpException)
{
    throw;  // Re-throw MCP exceptions as-is
}
catch (Exception ex)
{
    ThrowInternalError(ex, action, excelPath);  // Wrap with context
    throw;
}
```

### Helper Methods (ExcelToolsBase)
```csharp
// Throw for unknown action
ThrowUnknownAction(action, "list", "view", "import", ...)

// Throw for missing parameter
ThrowMissingParameter("queryName", action)

// Wrap exception with context
ThrowInternalError(ex, action, filePath)
```

---

## Async Handling

**Use `.GetAwaiter().GetResult()` for async Core methods:**

```csharp
private static string ImportPowerQuery(...)
{
    var result = commands.Import(excelPath, queryName, mCodeFile)
        .GetAwaiter().GetResult();  // Sync wrapper for async Core method
    
    return JsonSerializer.Serialize(result, JsonOptions);
}
```

**Why:** MCP tools are synchronous, but Core commands may be async.

---

## JSON Serialization

### Always Use JsonSerializer

```csharp
// ✅ CORRECT - Proper serialization
return JsonSerializer.Serialize(result, JsonOptions);

// ❌ WRONG - Windows path escaping issues
return $"{{ \"filePath\": \"{result.FilePath}\" }}";
```

### JSON Options
```csharp
private static readonly JsonSerializerOptions JsonOptions = new()
{
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    WriteIndented = false
};
```

---

## Security-First Design

### Power Query Privacy Levels
```csharp
// Optional parameter with clear documentation
string? privacyLevel = null  // None, Private, Organizational, Public
```

**Returns `PowerQueryPrivacyErrorResult` when needed but not specified:**
- Detected privacy levels from existing queries
- Recommended privacy level
- Explanation of implications
- Guidance on how to proceed

### VBA Trust
**Never modifies security settings automatically.**

Returns `VbaTrustRequiredResult` with manual setup instructions when VBA trust not enabled.

---

## MCP vs CLI Decision Matrix

| Scenario | Use MCP Server | Use CLI |
|----------|----------------|---------|
| **AI-assisted refactoring** | ✅ | ❌ |
| **Scripted workflows** | ❌ | ✅ |
| **Power Query optimization** | ✅ | ❌ |
| **VBA version control** | ❌ | ✅ |
| **Interactive code review** | ✅ | ❌ |
| **Automated testing** | ❌ | ✅ |

---

## server.json Synchronization

**⚠️ CRITICAL: Always update `src/ExcelMcp.McpServer/.mcp/server.json` when:**
- Adding new MCP tools
- Adding actions to existing tools
- Changing tool parameters
- Modifying tool descriptions

---

## Best Practices

1. **Throw McpException** - Don't return JSON errors
2. **Use async wrappers** - `.GetAwaiter().GetResult()` for Core async methods
3. **Validate parameters** - Use helper methods for clear errors
4. **Security defaults** - Never auto-apply sensitive settings
5. **Update server.json** - Keep configuration synchronized
6. **JSON serialization** - Always use `JsonSerializer`
7. **Clear error messages** - Include exception type, inner exceptions, context
8. **Handle JsonElement for complex parameters** - See "JSON Deserialization and COM Marshalling" section below

---

## JSON Deserialization and COM Marshalling

**⚠️ CRITICAL: MCP framework deserializes JSON arrays to `System.Text.Json.JsonElement`, NOT primitive types**

### The Problem

When MCP framework receives JSON parameters like `values: [["text", 123, true]]`, it deserializes to:
```csharp
List<List<object?>> values // Each object? is JsonElement, NOT string/int/bool!
```

Excel COM API requires proper C# types (string, int, double, bool). COM marshaller **CANNOT** convert `JsonElement` to Variant → Runtime exception.

### The Solution Pattern

For MCP tools accepting complex parameters (arrays, nested objects), add type conversion:

```csharp
private static object ConvertToCellValue(object? value)
{
    if (value == null)
        return string.Empty;

    // Handle System.Text.Json.JsonElement (from MCP JSON deserialization)
    if (value is System.Text.Json.JsonElement jsonElement)
    {
        return jsonElement.ValueKind switch
        {
            JsonValueKind.String => jsonElement.GetString() ?? string.Empty,
            JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => string.Empty,
            _ => jsonElement.ToString() ?? string.Empty
        };
    }

    // Already a proper type (from CLI or tests)
    return value;
}
```

### When to Apply This Pattern

**✅ REQUIRED for MCP tools accepting:**
- 2D arrays (`List<List<object?>>`) - like range values/formulas
- Object arrays with mixed types
- Nested JSON structures passed to COM APIs

**❌ NOT NEEDED for:**
- Simple string/int/bool parameters (MCP deserializes these correctly)
- Parameters not passed to Excel COM (internal processing only)

### Testing Strategy

Create integration tests simulating MCP JSON deserialization:
```csharp
[Fact]
public async Task MethodAsync_WithJsonElementValues_WorksCorrectly()
{
    // Simulate MCP framework JSON deserialization
    string json = """[["text", 123, true, null]]""";
    var jsonDoc = System.Text.Json.JsonDocument.Parse(json);
    
    // Convert to List<List<object?>> containing JsonElement objects
    var testData = new List<List<object?>>();
    foreach (var rowElement in jsonDoc.RootElement.EnumerateArray())
    {
        var row = new List<object?>();
        foreach (var cellElement in rowElement.EnumerateArray())
        {
            row.Add(cellElement); // This is JsonElement, not primitive!
        }
        testData.Add(row);
    }
    
    // Call method and verify it handles JsonElement correctly
    var result = await _commands.MethodAsync(batch, testData);
    Assert.True(result.Success);
}
```

### Real-World Bug Example

**Bug**: `excel_range` tool's `set-values` action failed with "Type 'System.Text.Json.JsonElement' cannot be marshalled to a Variant"

**Root Cause**: Direct assignment of `JsonElement` to `object[,]` array for COM interop

**Fix**: Added `ConvertToCellValue()` helper in `RangeCommands.Values.cs` to detect and convert `JsonElement` before COM assignment

**Why Tests Didn't Catch It**: Integration tests use C# literals (`new() { "test", 123 }`), CLI uses CSV parsing. Only MCP Server triggers JSON deserialization path.

---

## Modular Architecture

```
Tools/
├── ExcelToolsBase.cs        # Foundation utilities
├── ExcelFileTool.cs         # File operations
├── ExcelPowerQueryTool.cs   # Power Query M code
├── ExcelWorksheetTool.cs    # Sheet operations
├── ExcelParameterTool.cs    # Named ranges
├── ExcelCellTool.cs         # Cell operations
├── ExcelVbaTool.cs          # VBA macros
└── ExcelTools.cs            # Clean delegation
```

**Result:** 8 focused files instead of 649-line monolith, optimized for LLM understanding.

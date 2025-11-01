# MCP Tool Action Discoverability Issue

## Problem

**Current Architecture**: 10 tools with 95 total actions
- excel_powerquery: 12 actions (list, view, import, export, update, refresh, delete, set-load-to-table, etc.)
- excel_worksheet: 13 actions
- excel_range: 18 actions
- etc.

**Issue**: MCP clients see 10 tools but don't easily discover the 95 actions within them.
- Actions are documented in Description attributes
- Actions are validated via RegularExpression
- BUT: Clients don't expose this metadata in a user-friendly way

## Current Tool Design (Action-Based)

```csharp
[McpServerTool(Name = "excel_powerquery")]
[Description("...All actions: list, view, import, export...")]
public static async Task<string> ExcelPowerQuery(
    [RegularExpression("^(list|view|import|export|update|refresh|...)$")]
    string action,
    string excelPath,
    ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ListPowerQueries(...),
        "view" => ViewPowerQuery(...),
        ...
    };
}
```

**Pros:**
- ✅ Clean code organization (one class per domain)
- ✅ Shared validation and error handling
- ✅ Easy to maintain

**Cons:**
- ❌ Clients can't discover available actions easily
- ❌ Users must read documentation to know what actions exist
- ❌ No autocomplete for actions in most MCP clients

## Solution Options

### Option 1: Improve Action Discoverability (Keep Current Design)

**Add comprehensive prompts:**
```csharp
[McpServerPrompt(Name = "excel_powerquery_actions")]
public static string PowerQueryActions => @"
Available excel_powerquery actions:
- list: List all Power Queries
- view: View M code for a query
- import: Import query from .pq file
...
";
```

**Pros:**
- ✅ Minimal code changes
- ✅ Keeps clean architecture
- ✅ Prompts are discoverable via MCP protocol

**Cons:**
- ❌ Still requires users to read prompts
- ❌ Not all clients expose prompts well

### Option 2: Flatten to Individual Tools (MCP-Native)

**Create 95 separate tools:**
```csharp
[McpServerTool(Name = "excel_powerquery_list")]
public static Task<string> ListPowerQueries(string excelPath, string? batchId = null) => ...

[McpServerTool(Name = "excel_powerquery_import")]
public static Task<string> ImportPowerQuery(string excelPath, string queryName, string sourcePath, ...) => ...

[McpServerTool(Name = "excel_worksheet_create")]
public static Task<string> CreateWorksheet(string excelPath, string sheetName, ...) => ...
```

**Pros:**
- ✅ Perfect MCP protocol alignment
- ✅ Full tool discovery (clients see all 95 operations)
- ✅ Better autocomplete/suggestions
- ✅ More granular parameter validation (only relevant params per tool)

**Cons:**
- ❌ 95 tool registrations instead of 10
- ❌ More boilerplate code
- ❌ Tool list becomes very long

### Option 3: Hybrid Approach (RECOMMENDED)

**Keep action-based tools BUT add enum parameter with all actions:**

```csharp
public enum PowerQueryAction
{
    List,
    View,
    Import,
    Export,
    Update,
    Refresh,
    Delete,
    SetLoadToTable,
    SetLoadToDataModel,
    SetLoadToBoth,
    SetConnectionOnly,
    GetLoadConfig
}

[McpServerTool(Name = "excel_powerquery")]
public static async Task<string> ExcelPowerQuery(
    PowerQueryAction action,  // Enum instead of string
    string excelPath,
    ...)
{
    return action switch
    {
        PowerQueryAction.List => ...,
        PowerQueryAction.View => ...,
        ...
    };
}
```

**Pros:**
- ✅ MCP clients can introspect enum values (better discoverability)
- ✅ Type safety (no typos in action names)
- ✅ Keeps clean architecture
- ✅ Compile-time validation

**Cons:**
- ❌ Enums may not serialize well in MCP JSON-RPC
- ❌ Clients might see numeric values instead of names

## Recommendation

**Option 3 (Hybrid)** if MCP supports enum introspection  
**Option 1 (Prompts)** if enum introspection doesn't work  
**Option 2 (Flatten)** only if you want maximum MCP compliance

## Testing

To test which approach your MCP client prefers:
1. Check if client exposes parameter descriptions
2. Check if client exposes RegularExpression validation patterns
3. Check if client supports enum parameter types
4. Check if client shows MCP prompts

**Your client behavior will determine best solution.**

## Current Action Count by Tool

| Tool | Actions | Examples |
|------|---------|----------|
| excel_powerquery | 12 | list, view, import, refresh, set-load-to-data-model |
| excel_range | 18 | get-values, set-formulas, clear-contents, add-hyperlink |
| excel_worksheet | 13 | create, rename, delete, set-tab-color, hide |
| excel_datamodel | 13 | list-tables, create-measure, create-relationship |
| excel_connection | 8 | list, view, import, test, refresh |
| excel_table | 8 | list, create, resize, add-to-datamodel |
| excel_pivottable | 8 | create-from-range, add-row-field, refresh |
| excel_vba | 7 | list, view, import, run |
| excel_parameter | 5 | list, create, get, set |
| excel_file | 3 | create-empty, test |
| **TOTAL** | **95** | Across 10 tools |

## Implementation Priority

If you choose to improve discoverability:

1. **Add comprehensive MCP prompts** (1-2 hours)
   - One prompt per tool listing all actions with examples
   - Clients can query prompts to discover actions

2. **Improve tool descriptions** (30 mins)
   - Add formatted action lists to Description attributes
   - Use markdown formatting if supported

3. **Add action validation helper** (1 hour)
   - Create autocomplete helper that returns available actions
   - Useful for debugging and client development

4. **Consider enum approach** (2-3 hours)
   - Test if your MCP client supports enum introspection
   - If yes, convert action parameters to enums

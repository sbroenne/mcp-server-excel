# Suggested Next Actions - Design Document

> **Status**: Draft  
> **Author**: GitHub Copilot  
> **Date**: 2025-01-29  
> **Version**: 1.0

## Executive Summary

This document proposes a redesign of the `SuggestedNextActions` feature to address current brittleness, code duplication, and lack of type safety. The new design introduces a structured, context-aware system optimized for both MCP servers (LLM consumption) and CLI (human consumption).

## Problem Statement

### Current Issues

1. **String-based implementation** - All suggestions are hardcoded text scattered across ~40+ files
2. **Duplication** - Same logical actions have different string representations:
   - CLI: `"Use 'param-list' to see available parameters"`
   - MCP: `"Use 'list' to see available parameters"`
   - Core: Contains both variations depending on calling context
3. **No type safety** - Action names can contain typos, wrong casing, or reference non-existent actions
4. **Maintenance burden** - When adding/removing/renaming actions, must update suggestions in multiple places
5. **Inconsistent quality** - Some suggestions are helpful, others are generic or redundant
6. **Hardcoded validation** - Parameter validation rules duplicated across layers with no shared definitions

### Current Validation Implementation

The codebase has **three different validation approaches** with significant duplication:

#### 1. MCP Server - Data Annotation Attributes
```csharp
[McpServerTool(Name = "excel_powerquery")]
public static async Task<string> ExcelPowerQuery(
    [Required]
    [RegularExpression("^(list|view|import|export|update|refresh|delete|set-load-to-table|set-load-to-data-model|set-load-to-both|set-connection-only|get-load-config)$")]
    string action,
    
    [Required]
    [FileExtensions(Extensions = "xlsx,xlsm")]
    string excelPath,
    
    [StringLength(255, MinimumLength = 1)]
    string? queryName = null,
    
    [RegularExpression("^(None|Private|Organizational|Public)$")]
    string? privacyLevel = null)
```

**Problems:**
- Action list hardcoded in regex string (error-prone to maintain)
- File extensions duplicated across all 8 MCP tools
- Enum values hardcoded as strings (out of sync with actual enums)
- No compile-time validation of regex patterns

#### 2. CLI - Manual Argument Checking
```csharp
public int List(string[] args)
{
    if (args.Length < 2)
    {
        AnsiConsole.MarkupLine("[red]Usage:[/] pq-list <file.xlsx>");
        return 1;
    }
    // Manual validation continues...
}

public int View(string[] args)
{
    if (args.Length < 3)
    {
        AnsiConsole.MarkupLine("[red]Usage:[/] pq-view <file.xlsx> <query-name>");
        return 1;
    }
    // No validation of file extension, query name format, etc.
}
```

**Problems:**
- Each CLI command manually validates argument count
- No validation of argument content (file extensions, name formats, etc.)
- Usage messages hardcoded (can get out of sync with actual parameters)
- No reusable validation logic

#### 3. Core Commands - No Parameter Validation
```csharp
public async Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName)
{
    // NO validation of queryName format, length, null checks
    // Validation happens implicitly when Excel COM fails
}
```

**Problems:**
- Core layer assumes valid inputs from CLI/MCP
- Excel COM errors are cryptic (e.g., "Exception from HRESULT: 0x800A03EC")
- No early validation to provide helpful error messages

### Impact

- **Development friction** - Adding new commands requires updating suggestions in 3+ places
- **LLM confusion** - Inconsistent action names between MCP tools may confuse AI agents
- **User experience** - Generic suggestions like "Check file exists" don't add value
- **Regression risk** - Refactoring action names can break suggestions silently
- **Validation drift** - MCP, CLI, and Core can have different rules for same parameters
- **Maintenance burden** - Changing validation rules requires updates in multiple files
- **Poor error messages** - Users get COM errors instead of clear validation failures

## Design Principles

### For MCP Servers (LLM Consumption)

As an LLM working with an MCP server, I need:

1. **Action-based suggestions** - Tell me the exact MCP tool action to invoke next
   - ✅ Good: `{ tool: "excel_powerquery", action: "view", params: { queryName: "Sales" } }`
   - ❌ Bad: `"Use 'view' to inspect a query's M code"`

2. **Context-aware workflows** - Suggest logical next steps based on current state
   - After `list` → suggest `view`, `import`, or `delete`
   - After `create` with errors → suggest `list` to verify
   - After successful `update` → suggest `refresh` to reload data

3. **Parameter guidance** - Include required/optional parameter hints
   - `{ tool: "excel_powerquery", action: "view", requiredParams: ["queryName"] }`

4. **Error recovery** - Suggest specific diagnostic actions on failure
   - Privacy error → suggest privacy levels with explanations
   - VBA trust error → suggest trust setup action
   - Not found error → suggest `list` action to discover available items

5. **Workflow continuation** - Chain multi-step operations efficiently
   - Import query → Set load config → Refresh → Verify data
   - Create table → Apply filter → Sort → Export

### For CLI (Human Consumption)

For human CLI users, I need:

1. **Command examples** - Full command-line syntax with placeholders
   - ✅ Good: `excelcli pq-view workbook.xlsx "Sales"`
   - ❌ Bad: `Use 'view' action`

2. **Help text clarity** - Natural language explanations
   - ✅ Good: `"To see the M code for a query, use: pq-view <file> <query-name>"`
   - ❌ Bad: `"Action: view"`

3. **Next steps guidance** - What to do after successful/failed operations
   - After list → Show commands to view/import/delete
   - After error → Show commands to diagnose or fix

## Layer Separation Requirements

> **Critical Design Constraint**: Core layer MUST NOT contain presentation layer concerns

### Current Violations (See `layer-separation-analysis.md`)

The existing codebase violates separation of concerns:

1. **Core contains MCP/CLI-specific suggestions**
   - Core Commands populate `SuggestedNextActions` with MCP action names (`"Use 'view' to inspect..."`)
   - Core Commands also use CLI command names (`"Use 'param-list' to see..."`)
   - Core is tightly coupled to both presentation layers

2. **Mixed responsibility for suggestions**
   - Core Commands populate suggestions
   - MCP Tools sometimes overwrite, sometimes append
   - CLI just displays what Core provides
   - No clear ownership

3. **Validation scattered across layers**
   - MCP Server has data annotation attributes (only layer with validation!)
   - CLI does manual argument count checks (no content validation)
   - Core assumes inputs are valid (no parameter validation)
   - Rules can drift out of sync

### New Design Requirements

#### 1. Core Layer - Business Logic ONLY

```csharp
// Core/Models/ResultTypes.cs
public abstract class ResultBase
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? FilePath { get; set; }
    
    // ❌ REMOVE in v2.0:
    // public List<string> SuggestedNextActions { get; set; }
    // public string? WorkflowHint { get; set; }
}

// Core layer MUST NOT:
// - Know about CLI command names (pq-list, param-create, etc.)
// - Know about MCP action names (view, list, import, etc.)
// - Generate workflow guidance strings
// - Reference Spectre.Console, ModelContextProtocol, or any UI framework
```

**Rationale**: Core is business logic. Presentation is MCP/CLI responsibility.

#### 2. Shared Validation Layer

```csharp
// Core/Validation/ActionDefinitions.cs
public static class ActionDefinitions
{
    public static class PowerQuery
    {
        public static readonly ActionDefinition List = new(
            domain: "PowerQuery",
            action: "list",
            cliCommand: "pq-list",
            mcpAction: "list",
            mcpTool: "excel_powerquery",
            parameters: new[]
            {
                new ParameterDef("excelPath", required: true, 
                    fileExtensions: new[] { "xlsx", "xlsm" })
            }
        );
        
        public static readonly ActionDefinition View = new(
            domain: "PowerQuery",
            action: "view",
            cliCommand: "pq-view",
            mcpAction: "view",
            mcpTool: "excel_powerquery",
            parameters: new[]
            {
                new ParameterDef("excelPath", required: true, 
                    fileExtensions: new[] { "xlsx", "xlsm" }),
                new ParameterDef("queryName", required: true,
                    maxLength: 255, pattern: @"^[a-zA-Z0-9_]+$")
            }
        );
        
        // ... all actions defined once
    }
}

// Core/Validation/ParameterDefinition.cs
public class ParameterDef
{
    public string Name { get; }
    public bool Required { get; }
    public int? MinLength { get; }
    public int? MaxLength { get; }
    public string? Pattern { get; }  // Regex
    public string[]? FileExtensions { get; }
    public string[]? AllowedValues { get; }  // For enums
    public string? Description { get; }
    
    // Validation method
    public ValidationResult Validate(object? value) { ... }
}
```

**Benefits:**
- ✅ Single source of truth for all actions
- ✅ Validation rules shared between MCP and CLI
- ✅ CLI gets command name from definition
- ✅ MCP gets action name from definition
- ✅ Add action once, available everywhere
- ✅ Change action name once, propagates everywhere

#### 3. MCP Server - Uses ActionDefinitions

```csharp
// MCP/Tools/ExcelPowerQueryTool.cs
[McpServerTool(Name = "excel_powerquery")]
public static async Task<string> ExcelPowerQuery(
    [Required]
    [RegularExpression(ActionDefinitions.PowerQuery.GetActionRegex())]  // Generated from ActionDefinitions
    string action,
    
    [Required]
    [FileExtensions(Extensions = "xlsx,xlsm")]  // From ParameterDef
    string excelPath,
    
    [StringLength(255, MinimumLength = 1)]  // From ParameterDef
    string? queryName = null)
{
    // Validate using ActionDefinition
    var actionDef = ActionDefinitions.PowerQuery.GetByMcpAction(action);
    var validation = actionDef.ValidateParameters(new { excelPath, queryName });
    if (!validation.IsValid)
    {
        throw new McpException(validation.ErrorMessage);
    }
    
    // Call Core business logic
    var result = await commands.ViewAsync(batch, queryName);
    
    // Generate MCP-specific suggestions using ActionDefinition
    var nextActions = NextActionFactory.PowerQuery.GetNextActions(result, actionDef);
    
    return JsonSerializer.Serialize(new
    {
        result.Success,
        result.Data,
        NextActions = nextActions.Select(a => a.ToMcp()).ToList()
    });
}
```

#### 4. CLI - Uses ActionDefinitions

```csharp
// CLI/Commands/PowerQueryCommands.cs
public int View(string[] args)
{
    var actionDef = ActionDefinitions.PowerQuery.View;
    
    // Validate using ActionDefinition
    var validation = actionDef.ValidateCliArgs(args);
    if (!validation.IsValid)
    {
        AnsiConsole.MarkupLine($"[red]Error:[/] {validation.ErrorMessage}");
        AnsiConsole.MarkupLine($"[dim]Usage:[/] {actionDef.GetCliUsage()}");
        return 1;
    }
    
    // Call Core business logic
    var result = await commands.ViewAsync(batch, args[2]);
    
    // Generate CLI-specific suggestions using ActionDefinition
    if (result.Success)
    {
        DisplayResult(result);
        
        var nextActions = NextActionFactory.PowerQuery.GetNextActions(result, actionDef);
        DisplayNextActions(nextActions.Select(a => a.ToCli()));
    }
    
    return result.Success ? 0 : 1;
}
```

### Migration Path for Layer Separation

#### Phase 1: Add Shared Validation (Parallel to NextAction work)
1. Create `Core/Validation/ActionDefinitions.cs`
2. Create `Core/Validation/ParameterDefinition.cs`
3. Define all actions for PowerQuery, Parameter, Table, etc.
4. Write unit tests for validation logic

#### Phase 2: Update MCP Server
1. Generate regex patterns from ActionDefinitions
2. Use ParameterDef for attribute values
3. Validate using ActionDefinition.ValidateParameters()
4. Generate MCP suggestions from ActionDefinition

#### Phase 3: Update CLI
1. Use ActionDefinition.ValidateCliArgs()
2. Use ActionDefinition.GetCliUsage() for help text
3. Generate CLI suggestions from ActionDefinition

#### Phase 4: Clean Core (v2.0)
1. Remove `SuggestedNextActions` from ResultBase
2. Remove `WorkflowHint` from ResultBase
3. Remove all CLI command name references from Core
4. Remove all MCP action name references from Core
5. Core only returns business data

## Proposed Architecture

### 1. Core Abstraction - NextAction Types

```csharp
namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Represents a suggested next action that can be adapted for MCP or CLI context
/// </summary>
public abstract class NextAction
{
    /// <summary>
    /// Type of action for programmatic handling
    /// </summary>
    public abstract NextActionType ActionType { get; }
    
    /// <summary>
    /// Get MCP-formatted suggestion (tool name, action, parameters)
    /// </summary>
    public abstract NextActionMcp ToMcp();
    
    /// <summary>
    /// Get CLI-formatted suggestion (command example with syntax)
    /// </summary>
    public abstract NextActionCli ToCli();
    
    /// <summary>
    /// Get human-readable description
    /// </summary>
    public abstract string ToDescription();
}

/// <summary>
/// Categories of next actions
/// </summary>
public enum NextActionType
{
    /// <summary>View/inspect an item</summary>
    View,
    
    /// <summary>List available items</summary>
    List,
    
    /// <summary>Create new item</summary>
    Create,
    
    /// <summary>Update existing item</summary>
    Update,
    
    /// <summary>Delete item</summary>
    Delete,
    
    /// <summary>Refresh/reload data</summary>
    Refresh,
    
    /// <summary>Configure settings</summary>
    Configure,
    
    /// <summary>Verify/validate result</summary>
    Verify,
    
    /// <summary>Diagnose error</summary>
    Diagnose,
    
    /// <summary>Import from external source</summary>
    Import,
    
    /// <summary>Export to external destination</summary>
    Export,
    
    /// <summary>Execute/run operation</summary>
    Execute
}

/// <summary>
/// MCP-specific action representation (for LLM consumption)
/// </summary>
public class NextActionMcp
{
    /// <summary>Tool name (e.g., "excel_powerquery")</summary>
    public string Tool { get; init; } = "";
    
    /// <summary>Action within tool (e.g., "view")</summary>
    public string Action { get; init; } = "";
    
    /// <summary>Required parameters</summary>
    public Dictionary<string, string> RequiredParams { get; init; } = new();
    
    /// <summary>Optional parameters</summary>
    public Dictionary<string, string>? OptionalParams { get; init; }
    
    /// <summary>Brief rationale for this action</summary>
    public string? Rationale { get; init; }
}

/// <summary>
/// CLI-specific action representation (for human consumption)
/// </summary>
public class NextActionCli
{
    /// <summary>CLI command name (e.g., "pq-view")</summary>
    public string Command { get; init; } = "";
    
    /// <summary>Full command syntax example</summary>
    public string Example { get; init; } = "";
    
    /// <summary>Human-readable description</summary>
    public string Description { get; init; } = "";
}
```

### 2. Concrete Action Implementations

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// Suggests viewing an item (Power Query, VBA module, table, etc.)
/// </summary>
public class ViewItemAction : NextAction
{
    public override NextActionType ActionType => NextActionType.View;
    
    private readonly string _tool;
    private readonly string _cliCommand;
    private readonly string _itemType; // "query", "parameter", "table", etc.
    private readonly string? _itemName;
    private readonly string _filePlaceholder;
    
    public ViewItemAction(
        string tool, 
        string cliCommand, 
        string itemType, 
        string? itemName = null,
        string filePlaceholder = "<file>")
    {
        _tool = tool;
        _cliCommand = cliCommand;
        _itemType = itemType;
        _itemName = itemName;
        _filePlaceholder = filePlaceholder;
    }
    
    public override NextActionMcp ToMcp()
    {
        var required = new Dictionary<string, string>
        {
            ["excelPath"] = _filePlaceholder
        };
        
        if (_itemName != null)
        {
            required[$"{_itemType}Name"] = _itemName;
        }
        
        return new NextActionMcp
        {
            Tool = _tool,
            Action = "view",
            RequiredParams = required,
            Rationale = _itemName != null 
                ? $"Inspect details of {_itemType} '{_itemName}'"
                : $"View {_itemType} details"
        };
    }
    
    public override NextActionCli ToCli()
    {
        var example = _itemName != null
            ? $"excelcli {_cliCommand} {_filePlaceholder} \"{_itemName}\""
            : $"excelcli {_cliCommand} {_filePlaceholder} <{_itemType}-name>";
            
        return new NextActionCli
        {
            Command = _cliCommand,
            Example = example,
            Description = _itemName != null
                ? $"View details of {_itemType} '{_itemName}'"
                : $"View {_itemType} details"
        };
    }
    
    public override string ToDescription()
    {
        return _itemName != null
            ? $"View {_itemType} '{_itemName}'"
            : $"View {_itemType} details";
    }
}

/// <summary>
/// Suggests listing items in a workbook
/// </summary>
public class ListItemsAction : NextAction
{
    public override NextActionType ActionType => NextActionType.List;
    
    private readonly string _tool;
    private readonly string _cliCommand;
    private readonly string _itemTypePlural; // "queries", "parameters", "tables"
    
    public ListItemsAction(string tool, string cliCommand, string itemTypePlural)
    {
        _tool = tool;
        _cliCommand = cliCommand;
        _itemTypePlural = itemTypePlural;
    }
    
    public override NextActionMcp ToMcp()
    {
        return new NextActionMcp
        {
            Tool = _tool,
            Action = "list",
            RequiredParams = new Dictionary<string, string>
            {
                ["excelPath"] = "<file>"
            },
            Rationale = $"Discover available {_itemTypePlural}"
        };
    }
    
    public override NextActionCli ToCli()
    {
        return new NextActionCli
        {
            Command = _cliCommand,
            Example = $"excelcli {_cliCommand} <file>",
            Description = $"List all {_itemTypePlural} in workbook"
        };
    }
    
    public override string ToDescription() => $"List {_itemTypePlural}";
}

// Additional concrete classes for Create, Update, Delete, Configure, etc.
// Each follows same pattern: store context, implement ToMcp/ToCli/ToDescription
```

### 3. Context-Aware Factory

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// Factory for creating context-aware next action suggestions
/// </summary>
public static class NextActionFactory
{
    // PowerQuery Actions
    public static class PowerQuery
    {
        private const string Tool = "excel_powerquery";
        
        public static NextAction List() => 
            new ListItemsAction(Tool, "pq-list", "Power Queries");
            
        public static NextAction View(string queryName) =>
            new ViewItemAction(Tool, "pq-view", "query", queryName);
            
        public static NextAction Import(string queryName) =>
            new CreateItemAction(Tool, "pq-import", "query", queryName, 
                requiredFiles: new[] { "source.pq" });
                
        public static NextAction Update(string queryName) =>
            new UpdateItemAction(Tool, "pq-update", "query", queryName,
                requiredFiles: new[] { "source.pq" });
                
        public static NextAction Refresh(string queryName) =>
            new RefreshItemAction(Tool, "pq-refresh", "query", queryName);
                
        public static NextAction SetLoadToTable(string queryName, string? sheetName = null) =>
            new ConfigureAction(Tool, "pq-set-load-to-table", "query", queryName,
                optionalParams: sheetName != null 
                    ? new Dictionary<string, string> { ["targetSheet"] = sheetName }
                    : null);
                    
        public static NextAction CheckPrivacy() =>
            new DiagnoseAction(Tool, "pq-list", "privacy levels",
                rationale: "Privacy errors occur when combining data sources with different levels");
    }
    
    // Parameter Actions
    public static class Parameter
    {
        private const string Tool = "excel_parameter";
        
        public static NextAction List() =>
            new ListItemsAction(Tool, "param-list", "named parameters");
            
        public static NextAction Get(string paramName) =>
            new ViewItemAction(Tool, "param-get", "parameter", paramName);
            
        public static NextAction Set(string paramName, string value) =>
            new UpdateItemAction(Tool, "param-set", "parameter", paramName,
                additionalParams: new Dictionary<string, string> { ["value"] = value });
                
        public static NextAction Create(string paramName, string reference) =>
            new CreateItemAction(Tool, "param-create", "parameter", paramName,
                additionalParams: new Dictionary<string, string> { ["reference"] = reference });
    }
    
    // Table Actions
    public static class Table
    {
        private const string Tool = "excel_table";
        
        public static NextAction List() =>
            new ListItemsAction(Tool, "table-list", "Excel Tables");
            
        public static NextAction Info(string tableName) =>
            new ViewItemAction(Tool, "table-info", "table", tableName);
            
        public static NextAction Rename(string oldName, string newName) =>
            new UpdateItemAction(Tool, "table-rename", "table", oldName,
                additionalParams: new Dictionary<string, string> { ["newName"] = newName });
    }
    
    // ... Similar factories for VBA, DataModel, Worksheet, Cell, File, etc.
}
```

### 4. Result Type Integration

```csharp
namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Base result type updated to use NextAction abstraction
/// </summary>
public abstract class ResultBase
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? FilePath { get; set; }
    
    /// <summary>
    /// Structured next actions (replaces string-based SuggestedNextActions)
    /// </summary>
    public List<NextAction> NextActions { get; set; } = new();
    
    /// <summary>
    /// Backward compatibility - generates strings from NextActions
    /// </summary>
    [Obsolete("Use NextActions instead. This property will be removed in v2.0")]
    public List<string> SuggestedNextActions 
    {
        get => NextActions.Select(a => a.ToDescription()).ToList();
        set { } // Ignore sets for backward compatibility
    }
    
    /// <summary>
    /// Contextual workflow hint for LLM
    /// </summary>
    public string? WorkflowHint { get; set; }
}
```

### 5. Usage in Commands

```csharp
// BEFORE (PowerQueryCommands.cs)
result.SuggestedNextActions = new List<string>
{
    "Use 'view' to inspect a query's M code",
    "Use 'import' to add a new Power Query",
    "Use 'delete' to remove a query"
};

// AFTER
result.NextActions = new List<NextAction>
{
    NextActionFactory.PowerQuery.View("<query-name>"),
    NextActionFactory.PowerQuery.Import("<query-name>"),
    NextActionFactory.PowerQuery.Delete("<query-name>")
};
```

### 6. Usage in MCP Tools

```csharp
// ExcelPowerQueryTool.cs

private static async Task<string> ListPowerQueriesAsync(...)
{
    var result = await commands.ListAsync(batch);
    
    if (!result.Success)
    {
        result.NextActions = new List<NextAction>
        {
            NextActionFactory.PowerQuery.CheckFileExists(),
            NextActionFactory.PowerQuery.VerifyFilePath()
        };
        throw new McpException(...);
    }
    
    // Context-aware suggestions based on result
    if (result.Queries.Any())
    {
        var firstQuery = result.Queries.First().Name;
        result.NextActions = new List<NextAction>
        {
            NextActionFactory.PowerQuery.View(firstQuery),
            NextActionFactory.PowerQuery.Import("<new-query>"),
            NextActionFactory.PowerQuery.Delete("<query-name>")
        };
    }
    else
    {
        result.NextActions = new List<NextAction>
        {
            NextActionFactory.PowerQuery.Import("<query-name>")
        };
    }
    
    // Serialize with MCP format
    var mcpResult = new
    {
        result.Success,
        result.Queries,
        NextActions = result.NextActions.Select(a => a.ToMcp()).ToList()
    };
    
    return JsonSerializer.Serialize(mcpResult, ExcelToolsBase.JsonOptions);
}
```

### 7. Usage in CLI Commands

```csharp
// CLI/Commands/PowerQueryCommands.cs

public int List(string[] args)
{
    var result = await _coreCommands.ListAsync(batch);
    
    if (result.Success)
    {
        // Display results...
        
        if (result.NextActions.Any())
        {
            AnsiConsole.MarkupLine("\n[bold]Suggested Next Steps:[/]");
            foreach (var action in result.NextActions)
            {
                var cli = action.ToCli();
                AnsiConsole.MarkupLine($"  • [cyan]{cli.Description}[/]");
                AnsiConsole.MarkupLine($"    {cli.Example.EscapeMarkup()}");
            }
        }
    }
    
    return result.Success ? 0 : 1;
}
```

## Implementation Strategy

### Phase 1: Core Infrastructure (Week 1)
1. Create `NextAction` base class and concrete implementations
2. Create `NextActionFactory` with all action builders
3. Add `NextActions` property to `ResultBase`
4. Mark `SuggestedNextActions` as obsolete (but maintain for backward compatibility)
5. Write unit tests for action serialization

### Phase 2: Core Commands Migration (Week 2)
1. Update `PowerQueryCommands` to use `NextActionFactory`
2. Update `ParameterCommands` to use `NextActionFactory`
3. Update `TableCommands` to use `NextActionFactory`
4. Update `DataModelCommands` to use `NextActionFactory`
5. Update `ScriptCommands` (VBA) to use `NextActionFactory`
6. Write integration tests

### Phase 3: MCP Server Migration (Week 3)
1. Update `ExcelPowerQueryTool` to serialize `NextActions.ToMcp()`
2. Update `ExcelParameterTool` to serialize `NextActions.ToMcp()`
3. Update `TableTool` to serialize `NextActions.ToMcp()`
4. Update `ExcelDataModelTool` to serialize `NextActions.ToMcp()`
5. Update `ExcelVbaTool` to serialize `NextActions.ToMcp()`
6. Write MCP integration tests

### Phase 4: CLI Migration (Week 4)
1. Update CLI `PowerQueryCommands` to display `NextActions.ToCli()`
2. Update CLI `ParameterCommands` to display `NextActions.ToCli()`
3. Update CLI `TableCommands` to display `NextActions.ToCli()`
4. Update CLI `DataModelCommands` to display `NextActions.ToCli()`
5. Update CLI `ScriptCommands` to display `NextActions.ToCli()`
6. Write CLI integration tests

### Phase 5: Deprecation (v2.0)
1. Remove `SuggestedNextActions` property
2. Update documentation
3. Release major version

## Testing Strategy

### Unit Tests
```csharp
[Fact]
public void ViewItemAction_ToMcp_ReturnsCorrectFormat()
{
    var action = new ViewItemAction("excel_powerquery", "pq-view", "query", "Sales");
    var mcp = action.ToMcp();
    
    Assert.Equal("excel_powerquery", mcp.Tool);
    Assert.Equal("view", mcp.Action);
    Assert.Equal("Sales", mcp.RequiredParams["queryName"]);
}

[Fact]
public void ViewItemAction_ToCli_ReturnsCorrectExample()
{
    var action = new ViewItemAction("excel_powerquery", "pq-view", "query", "Sales");
    var cli = action.ToCli();
    
    Assert.Equal("pq-view", cli.Command);
    Assert.Contains("pq-view", cli.Example);
    Assert.Contains("Sales", cli.Example);
}
```

### Integration Tests
```csharp
[Fact]
public async Task ListAsync_ReturnsViewActionForFirstQuery()
{
    var result = await _commands.ListAsync(batch);
    
    Assert.True(result.NextActions.Any(a => a.ActionType == NextActionType.View));
    
    var viewAction = result.NextActions.First(a => a.ActionType == NextActionType.View);
    var mcp = viewAction.ToMcp();
    Assert.Equal("excel_powerquery", mcp.Tool);
    Assert.Equal("view", mcp.Action);
}
```

## Benefits

### For Development
✅ Type-safe action references (compiler catches typos)  
✅ Single source of truth for each action  
✅ Easy to add new actions (one factory method)  
✅ Refactoring-friendly (rename action, all suggestions update)  
✅ Testable (unit test each action type)

### For LLMs (MCP)
✅ Structured JSON with tool/action/params  
✅ Required vs optional parameters clearly marked  
✅ Rationale explains why action is suggested  
✅ Context-aware (different suggestions based on state)  
✅ Workflow chaining (multi-step operations)

### For Humans (CLI)
✅ Full command examples with placeholders  
✅ Natural language descriptions  
✅ Copy-paste ready syntax  
✅ Consistent formatting across all commands

## Migration Path

### Backward Compatibility
The `SuggestedNextActions` property remains functional during migration:
- Core commands populate both `NextActions` and `SuggestedNextActions`
- MCP tools read `NextActions` and serialize to MCP format
- CLI tools read `NextActions` and serialize to CLI format
- Old code reading `SuggestedNextActions` still works

### Breaking Changes (v2.0)
- Remove `SuggestedNextActions` property
- Remove backward compatibility shims
- Require all consumers to use `NextActions`

## Example Workflow

### LLM Agent Using MCP Server

```json
// User: "List Power Queries in workbook.xlsx"

// LLM calls:
{
  "tool": "excel_powerquery",
  "action": "list",
  "params": { "excelPath": "workbook.xlsx" }
}

// MCP responds:
{
  "success": true,
  "queries": [
    { "name": "Sales", "isConnectionOnly": false },
    { "name": "Customers", "isConnectionOnly": true }
  ],
  "nextActions": [
    {
      "tool": "excel_powerquery",
      "action": "view",
      "requiredParams": { "excelPath": "workbook.xlsx", "queryName": "Sales" },
      "rationale": "Inspect M code of 'Sales' query"
    },
    {
      "tool": "excel_powerquery",
      "action": "set-load-to-table",
      "requiredParams": { "excelPath": "workbook.xlsx", "queryName": "Customers" },
      "rationale": "'Customers' is connection-only, load to worksheet to validate"
    }
  ]
}

// LLM: "Found 2 queries. Let me view the Sales query to understand its logic."
// LLM calls:
{
  "tool": "excel_powerquery",
  "action": "view",
  "params": { "excelPath": "workbook.xlsx", "queryName": "Sales" }
}
```

### Human User Using CLI

```bash
$ excelcli pq-list workbook.xlsx

Power Queries:
  1. Sales (Loaded to Sheet1)
  2. Customers (Connection only)

Suggested Next Steps:
  • View query 'Sales' details
    excelcli pq-view workbook.xlsx "Sales"
  • Load 'Customers' to worksheet for validation
    excelcli pq-set-load-to-table workbook.xlsx "Customers" "Customers"

$ excelcli pq-view workbook.xlsx "Sales"

Query: Sales
M Code:
  let
    Source = Excel.CurrentWorkbook(){[Name="SalesData"]}[Content],
    ...
  in
    Result

Suggested Next Steps:
  • Update query M code
    excelcli pq-update workbook.xlsx "Sales" source.pq
  • Refresh query data
    excelcli pq-refresh workbook.xlsx "Sales"
```

## Open Questions

1. **Parameter formatting** - Should examples use `<placeholder>` or `{variable}` syntax?
   - Recommendation: `<placeholder>` for readability

2. **Action priorities** - Should `NextActions` be ordered by priority?
   - Recommendation: Yes, most common/useful actions first

3. **Context limits** - How many next actions to suggest?
   - Recommendation: 3-5 for MCP (LLM can process all), 2-3 for CLI (human attention span)

4. **Dynamic parameters** - Should actions include actual values from result?
   - Recommendation: Yes when available (e.g., first query name after list)

5. **Localization** - Should CLI descriptions support multiple languages?
   - Recommendation: Defer to v2.0, English-only for v1.0

## Conclusion

This design addresses all current issues while providing a foundation for future enhancements:
- **Type safety** eliminates brittle string manipulation
- **Context awareness** provides smarter suggestions
- **Dual format** serves both LLM and human users optimally
- **Factory pattern** centralizes action definitions
- **Backward compatibility** enables gradual migration

The investment in this refactoring will pay dividends through:
- Reduced maintenance burden
- Improved LLM agent effectiveness
- Better user experience for CLI users
- Easier addition of new features

## References

- [Model Context Protocol Specification](https://modelcontextprotocol.io/)
- [GitHub Copilot Best Practices](https://docs.github.com/en/copilot)
- [Command Pattern (Gang of Four)](https://refactoring.guru/design-patterns/command)

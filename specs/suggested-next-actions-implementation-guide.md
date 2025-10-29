# SuggestedNextActions - Implementation Guide

> **Companion to**: `suggested-next-actions-design.md`  
> **Purpose**: Step-by-step implementation instructions  
> **Audience**: Developers implementing the design

## Quick Start

This guide provides concrete implementation steps for the SuggestedNextActions refactoring described in the design document.

## File Organization

### New Files to Create

```
src/ExcelMcp.Core/Models/
├── NextActions/
│   ├── NextAction.cs              # Base abstract class
│   ├── NextActionType.cs          # Enum of action types
│   ├── NextActionMcp.cs           # MCP format representation
│   ├── NextActionCli.cs           # CLI format representation
│   ├── ViewItemAction.cs          # Concrete: View item
│   ├── ListItemsAction.cs         # Concrete: List items
│   ├── CreateItemAction.cs        # Concrete: Create item
│   ├── UpdateItemAction.cs        # Concrete: Update item
│   ├── DeleteItemAction.cs        # Concrete: Delete item
│   ├── RefreshItemAction.cs       # Concrete: Refresh/reload
│   ├── ConfigureAction.cs         # Concrete: Configure settings
│   ├── DiagnoseAction.cs          # Concrete: Diagnose errors
│   ├── ImportAction.cs            # Concrete: Import from file
│   ├── ExportAction.cs            # Concrete: Export to file
│   └── NextActionFactory.cs       # Factory with domain builders
└── Validation/                    # NEW - Shared validation layer
    ├── ActionDefinition.cs        # Action metadata (CLI + MCP)
    ├── ParameterDefinition.cs     # Parameter validation rules
    ├── ValidationResult.cs        # Validation result type
    └── ActionDefinitions.cs       # Central action registry
```

### Files to Modify

```
src/ExcelMcp.Core/Models/
└── ResultTypes.cs                  # Add NextActions property, mark SuggestedNextActions obsolete

src/ExcelMcp.Core/Commands/
├── PowerQueryCommands.cs           # Update to use NextActionFactory
├── ParameterCommands.cs            # Update to use NextActionFactory
├── ScriptCommands.cs               # Update to use NextActionFactory
├── DataModel/DataModelCommands.*.cs # Update to use NextActionFactory
└── Table/TableCommands.*.cs        # Update to use NextActionFactory

src/ExcelMcp.McpServer/Tools/
├── ExcelPowerQueryTool.cs          # Serialize NextActions.ToMcp()
├── ExcelParameterTool.cs           # Serialize NextActions.ToMcp()
├── ExcelVbaTool.cs                 # Serialize NextActions.ToMcp()
├── ExcelDataModelTool.cs           # Serialize NextActions.ToMcp()
└── TableTool.cs                    # Serialize NextActions.ToMcp()

src/ExcelMcp.CLI/Commands/
├── PowerQueryCommands.cs           # Display NextActions.ToCli()
├── ParameterCommands.cs            # Display NextActions.ToCli()
├── ScriptCommands.cs               # Display NextActions.ToCli()
├── DataModelCommands.cs            # Display NextActions.ToCli()
└── TableCommands.cs                # Display NextActions.ToCli()
```

## Step-by-Step Implementation

### Phase 1: Core Abstractions

#### Step 1.1: Create Base Types

Create `src/ExcelMcp.Core/Models/NextActions/NextActionType.cs`:

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// Categories of next actions for workflow guidance
/// </summary>
public enum NextActionType
{
    /// <summary>View/inspect an item (query, parameter, table, etc.)</summary>
    View,
    
    /// <summary>List available items in workbook</summary>
    List,
    
    /// <summary>Create new item</summary>
    Create,
    
    /// <summary>Update existing item</summary>
    Update,
    
    /// <summary>Delete item</summary>
    Delete,
    
    /// <summary>Refresh/reload data from source</summary>
    Refresh,
    
    /// <summary>Configure settings or properties</summary>
    Configure,
    
    /// <summary>Verify/validate operation result</summary>
    Verify,
    
    /// <summary>Diagnose error or problem</summary>
    Diagnose,
    
    /// <summary>Import from external source</summary>
    Import,
    
    /// <summary>Export to external destination</summary>
    Export,
    
    /// <summary>Execute/run operation</summary>
    Execute
}
```

Create `src/ExcelMcp.Core/Models/NextActions/NextActionMcp.cs`:

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// MCP-specific action representation for LLM consumption
/// Provides structured JSON format with tool, action, and parameters
/// </summary>
public class NextActionMcp
{
    /// <summary>
    /// MCP tool name (e.g., "excel_powerquery", "excel_parameter")
    /// </summary>
    public string Tool { get; init; } = "";
    
    /// <summary>
    /// Action within the tool (e.g., "view", "list", "create")
    /// </summary>
    public string Action { get; init; } = "";
    
    /// <summary>
    /// Required parameters for this action
    /// Key: parameter name, Value: example or placeholder value
    /// </summary>
    public Dictionary<string, string> RequiredParams { get; init; } = new();
    
    /// <summary>
    /// Optional parameters for this action
    /// Key: parameter name, Value: example or default value
    /// </summary>
    public Dictionary<string, string>? OptionalParams { get; init; }
    
    /// <summary>
    /// Brief rationale explaining why this action is suggested
    /// Helps LLM understand the workflow context
    /// </summary>
    public string? Rationale { get; init; }
}
```

Create `src/ExcelMcp.Core/Models/NextActions/NextActionCli.cs`:

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// CLI-specific action representation for human consumption
/// Provides command-line syntax examples and descriptions
/// </summary>
public class NextActionCli
{
    /// <summary>
    /// CLI command name (e.g., "pq-view", "param-list")
    /// </summary>
    public string Command { get; init; } = "";
    
    /// <summary>
    /// Full command syntax example with placeholders
    /// Example: "excelcli pq-view &lt;file&gt; &lt;query-name&gt;"
    /// </summary>
    public string Example { get; init; } = "";
    
    /// <summary>
    /// Human-readable description of what the command does
    /// </summary>
    public string Description { get; init; } = "";
}
```

Create `src/ExcelMcp.Core/Models/NextActions/NextAction.cs`:

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// Base class for suggested next actions
/// Provides dual-format output for both MCP (LLM) and CLI (human) contexts
/// </summary>
public abstract class NextAction
{
    /// <summary>
    /// Type of action for categorization and filtering
    /// </summary>
    public abstract NextActionType ActionType { get; }
    
    /// <summary>
    /// Convert to MCP format for LLM consumption
    /// Includes structured parameters and rationale
    /// </summary>
    public abstract NextActionMcp ToMcp();
    
    /// <summary>
    /// Convert to CLI format for human consumption
    /// Includes command syntax examples
    /// </summary>
    public abstract NextActionCli ToCli();
    
    /// <summary>
    /// Get human-readable description
    /// Used for backward compatibility and simple display
    /// </summary>
    public abstract string ToDescription();
}
```

#### Step 1.2: Update ResultBase

Modify `src/ExcelMcp.Core/Models/ResultTypes.cs`:

```csharp
// At the top of ResultBase class, add:

/// <summary>
/// Structured next actions (NEW - recommended)
/// Provides context-aware suggestions in both MCP and CLI formats
/// </summary>
public List<NextAction> NextActions { get; set; } = new();

/// <summary>
/// Legacy string-based suggestions (DEPRECATED)
/// Use NextActions instead - will be removed in v2.0
/// </summary>
[Obsolete("Use NextActions instead. This property generates strings from NextActions for backward compatibility.")]
public List<string> SuggestedNextActions 
{
    get => NextActions.Select(a => a.ToDescription()).ToList();
    set { /* Ignore sets - backward compatibility only */ }
}
```

### Phase 2: Concrete Action Implementations

#### Step 2.1: Create Common Actions

Create `src/ExcelMcp.Core/Models/NextActions/ViewItemAction.cs`:

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// Suggests viewing details of a specific item (query, parameter, table, etc.)
/// </summary>
public class ViewItemAction : NextAction
{
    public override NextActionType ActionType => NextActionType.View;
    
    private readonly string _tool;
    private readonly string _cliCommand;
    private readonly string _itemType;
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
```

Create similar implementations for:
- `ListItemsAction.cs` - List all items of a type
- `CreateItemAction.cs` - Create new item with parameters
- `UpdateItemAction.cs` - Update existing item
- `DeleteItemAction.cs` - Delete item
- `RefreshItemAction.cs` - Refresh/reload data
- `ConfigureAction.cs` - Configure settings
- `DiagnoseAction.cs` - Diagnose errors

#### Step 2.2: Create Factory

Create `src/ExcelMcp.Core/Models/NextActions/NextActionFactory.cs`:

```csharp
namespace Sbroenne.ExcelMcp.Core.Models.NextActions;

/// <summary>
/// Factory for creating context-aware next action suggestions
/// Organizes actions by domain (PowerQuery, Parameter, Table, etc.)
/// </summary>
public static class NextActionFactory
{
    /// <summary>
    /// Power Query related actions
    /// </summary>
    public static class PowerQuery
    {
        private const string Tool = "excel_powerquery";
        
        public static NextAction List() => 
            new ListItemsAction(Tool, "pq-list", "Power Queries");
            
        public static NextAction View(string queryName) =>
            new ViewItemAction(Tool, "pq-view", "query", queryName);
            
        public static NextAction Import(string? queryName = null) =>
            new ImportAction(Tool, "pq-import", "query", queryName ?? "<query-name>",
                requiredFiles: new[] { "source.pq" });
                
        public static NextAction Update(string queryName) =>
            new UpdateItemAction(Tool, "pq-update", "query", queryName,
                requiredFiles: new[] { "source.pq" });
                
        public static NextAction Delete(string queryName) =>
            new DeleteItemAction(Tool, "pq-delete", "query", queryName);
                
        public static NextAction Refresh(string queryName) =>
            new RefreshItemAction(Tool, "pq-refresh", "query", queryName);
                
        public static NextAction SetLoadToTable(string queryName, string? sheetName = null) =>
            new ConfigureAction(Tool, "pq-set-load-to-table", "query load", queryName,
                additionalParams: sheetName != null 
                    ? new Dictionary<string, string> { ["targetSheet"] = sheetName }
                    : null);
    }
    
    /// <summary>
    /// Named Parameter related actions
    /// </summary>
    public static class Parameter
    {
        private const string Tool = "excel_parameter";
        
        public static NextAction List() =>
            new ListItemsAction(Tool, "param-list", "named parameters");
            
        public static NextAction Get(string paramName) =>
            new ViewItemAction(Tool, "param-get", "parameter", paramName);
            
        public static NextAction Set(string paramName, string? value = null) =>
            new UpdateItemAction(Tool, "param-set", "parameter", paramName,
                additionalParams: value != null 
                    ? new Dictionary<string, string> { ["value"] = value }
                    : null);
                
        public static NextAction Create(string paramName, string? reference = null) =>
            new CreateItemAction(Tool, "param-create", "parameter", paramName,
                additionalParams: reference != null
                    ? new Dictionary<string, string> { ["reference"] = reference }
                    : null);
                    
        public static NextAction Delete(string paramName) =>
            new DeleteItemAction(Tool, "param-delete", "parameter", paramName);
    }
    
    // TODO: Add Table, VBA, DataModel, Worksheet, Range factories
}
```

### Phase 3: Migration Examples

#### Example 1: PowerQueryCommands.ListAsync

**BEFORE:**
```csharp
result.SuggestedNextActions = new List<string>
{
    "Use 'view' to inspect a query's M code",
    "Use 'import' to add a new Power Query",
    "Use 'delete' to remove a query"
};
```

**AFTER:**
```csharp
result.NextActions = new List<NextAction>();

if (result.Queries.Any())
{
    var firstQuery = result.Queries.First().Name;
    result.NextActions.Add(NextActionFactory.PowerQuery.View(firstQuery));
    result.NextActions.Add(NextActionFactory.PowerQuery.Import());
    result.NextActions.Add(NextActionFactory.PowerQuery.Delete("<query-name>"));
}
else
{
    result.NextActions.Add(NextActionFactory.PowerQuery.Import());
}
```

#### Example 2: ExcelPowerQueryTool (MCP)

**BEFORE:**
```csharp
result.SuggestedNextActions = new List<string>
{
    "Use 'view' to inspect a query's M code",
    "Use 'import' to add a new Power Query"
};
result.WorkflowHint = "Power Queries listed...";

return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
```

**AFTER:**
```csharp
result.NextActions = new List<NextAction>
{
    NextActionFactory.PowerQuery.View("<query-name>"),
    NextActionFactory.PowerQuery.Import()
};
result.WorkflowHint = "Power Queries listed. View, import, or delete as needed.";

// Serialize with MCP format
var mcpResult = new
{
    result.Success,
    result.Queries,
    NextActions = result.NextActions.Select(a => a.ToMcp()).ToList(),
    result.WorkflowHint
};

return JsonSerializer.Serialize(mcpResult, ExcelToolsBase.JsonOptions);
```

#### Example 3: CLI PowerQueryCommands

**BEFORE:**
```csharp
if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
{
    AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
    foreach (var suggestion in result.SuggestedNextActions)
    {
        AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
    }
}
```

**AFTER:**
```csharp
if (result.NextActions != null && result.NextActions.Any())
{
    AnsiConsole.MarkupLine("\n[bold]Suggested Next Steps:[/]");
    foreach (var action in result.NextActions)
    {
        var cli = action.ToCli();
        AnsiConsole.MarkupLine($"  • [cyan]{cli.Description}[/]");
        AnsiConsole.MarkupLine($"    [dim]{cli.Example.EscapeMarkup()}[/]");
    }
}
```

## Testing Strategy

### Unit Tests

Create `tests/ExcelMcp.Core.Tests/Unit/NextActions/NextActionTests.cs`:

```csharp
using Sbroenne.ExcelMcp.Core.Models.NextActions;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit.NextActions;

public class NextActionTests
{
    [Fact]
    public void ViewItemAction_ToMcp_ReturnsCorrectToolAndAction()
    {
        var action = new ViewItemAction("excel_powerquery", "pq-view", "query", "Sales");
        var mcp = action.ToMcp();
        
        Assert.Equal("excel_powerquery", mcp.Tool);
        Assert.Equal("view", mcp.Action);
        Assert.Equal("Sales", mcp.RequiredParams["queryName"]);
        Assert.Contains("Sales", mcp.Rationale);
    }
    
    [Fact]
    public void ViewItemAction_ToCli_ReturnsCorrectExample()
    {
        var action = new ViewItemAction("excel_powerquery", "pq-view", "query", "Sales");
        var cli = action.ToCli();
        
        Assert.Equal("pq-view", cli.Command);
        Assert.Contains("pq-view", cli.Example);
        Assert.Contains("Sales", cli.Example);
        Assert.Contains("query", cli.Description.ToLower());
    }
    
    [Fact]
    public void ListItemsAction_ToMcp_ReturnsListAction()
    {
        var action = new ListItemsAction("excel_powerquery", "pq-list", "queries");
        var mcp = action.ToMcp();
        
        Assert.Equal("excel_powerquery", mcp.Tool);
        Assert.Equal("list", mcp.Action);
        Assert.Contains("excelPath", mcp.RequiredParams.Keys);
    }
    
    [Fact]
    public void NextActionFactory_PowerQuery_CreatesCorrectActions()
    {
        var viewAction = NextActionFactory.PowerQuery.View("Sales");
        Assert.Equal(NextActionType.View, viewAction.ActionType);
        
        var listAction = NextActionFactory.PowerQuery.List();
        Assert.Equal(NextActionType.List, listAction.ActionType);
        
        var importAction = NextActionFactory.PowerQuery.Import();
        Assert.Equal(NextActionType.Import, importAction.ActionType);
    }
}
```

### Integration Tests

Create tests verifying that:
1. Core commands populate `NextActions` correctly
2. MCP tools serialize `NextActions.ToMcp()` as valid JSON
3. CLI tools display `NextActions.ToCli()` with proper formatting

## Migration Checklist

### Phase 1: Core Infrastructure
- [ ] Create `NextActionType.cs`
- [ ] Create `NextActionMcp.cs`
- [ ] Create `NextActionCli.cs`
- [ ] Create `NextAction.cs` base class
- [ ] Update `ResultTypes.cs` with `NextActions` property
- [ ] Mark `SuggestedNextActions` as `[Obsolete]`
- [ ] Create concrete action classes (View, List, Create, etc.)
- [ ] Create `NextActionFactory.cs` with PowerQuery/Parameter sections
- [ ] Write unit tests for actions

### Phase 2: Core Commands
- [ ] Update `PowerQueryCommands.cs` (List, View, Import, etc.)
- [ ] Update `ParameterCommands.cs` (List, Get, Set, Create, Delete)
- [ ] Update `ScriptCommands.cs` (VBA operations)
- [ ] Update `TableCommands.*.cs` (Table lifecycle, filters, sort)
- [ ] Update `DataModelCommands.*.cs` (Measures, relationships, etc.)
- [ ] Write integration tests

### Phase 3: MCP Server
- [ ] Update `ExcelPowerQueryTool.cs` serialization
- [ ] Update `ExcelParameterTool.cs` serialization
- [ ] Update `ExcelVbaTool.cs` serialization
- [ ] Update `TableTool.cs` serialization
- [ ] Update `ExcelDataModelTool.cs` serialization
- [ ] Write MCP integration tests

### Phase 4: CLI
- [ ] Update CLI `PowerQueryCommands.cs` display logic
- [ ] Update CLI `ParameterCommands.cs` display logic
- [ ] Update CLI `ScriptCommands.cs` display logic
- [ ] Update CLI `TableCommands.cs` display logic
- [ ] Update CLI `DataModelCommands.cs` display logic
- [ ] Write CLI integration tests

### Phase 5: Documentation
- [ ] Update README with NextActions examples
- [ ] Update API documentation
- [ ] Add migration guide for external consumers
- [ ] Mark deprecation timeline for `SuggestedNextActions`

## Common Patterns

### Pattern 1: Context-Aware Suggestions After List

```csharp
// Core command
public async Task<PowerQueryListResult> ListAsync(ExcelBatch batch)
{
    var result = new PowerQueryListResult { FilePath = batch.FilePath };
    
    // ... populate result.Queries ...
    
    // Context-aware next actions
    if (result.Queries.Any())
    {
        var firstQuery = result.Queries.First().Name;
        result.NextActions.Add(NextActionFactory.PowerQuery.View(firstQuery));
        result.NextActions.Add(NextActionFactory.PowerQuery.Import());
        result.NextActions.Add(NextActionFactory.PowerQuery.Delete("<query-name>"));
    }
    else
    {
        result.NextActions.Add(NextActionFactory.PowerQuery.Import());
    }
    
    return result;
}
```

### Pattern 2: Error Recovery Suggestions

```csharp
// When operation fails
catch (Exception ex)
{
    result.Success = false;
    result.ErrorMessage = ex.Message;
    result.NextActions.Add(NextActionFactory.PowerQuery.List());
    result.NextActions.Add(NextActionFactory.Common.VerifyFileExists());
    return result;
}
```

### Pattern 3: Workflow Continuation

```csharp
// After successful create
result.NextActions.Add(NextActionFactory.PowerQuery.SetLoadToTable(queryName));
result.NextActions.Add(NextActionFactory.PowerQuery.Refresh(queryName));
result.NextActions.Add(NextActionFactory.PowerQuery.View(queryName));
```

## Success Criteria

✅ **Type Safety**: No compilation with invalid action references  
✅ **DRY**: Each action defined once in factory  
✅ **Context-Aware**: Different suggestions based on operation result  
✅ **Dual Format**: Both MCP and CLI formats work correctly  
✅ **Backward Compatible**: Old code using `SuggestedNextActions` still works  
✅ **Tested**: 80%+ coverage for new action classes  
✅ **Documented**: Clear examples in docs and XML comments

## Timeline Estimate

- **Phase 1** (Core Infrastructure): 1-2 days
- **Phase 2** (Core Commands): 2-3 days
- **Phase 3** (MCP Server): 1-2 days
- **Phase 4** (CLI): 1-2 days
- **Phase 5** (Documentation): 1 day

**Total**: 6-10 days for complete migration

## Questions / Issues

Track implementation questions here:

1. **Q**: Should we validate action parameters in factory methods?
   **A**: TBD - probably yes for required params

2. **Q**: How to handle file path placeholders in different contexts?
   **A**: TBD - use actual file path when available, `<file>` otherwise

3. **Q**: Should NextActions be serialized in all result types?
   **A**: TBD - yes for MCP, optional for CLI (can use ToDescription())

## References

- Design Document: `specs/suggested-next-actions-design.md`
- MCP Specification: https://modelcontextprotocol.io/
- GitHub Issue: [FEATURE] Review suggested next actions set-up

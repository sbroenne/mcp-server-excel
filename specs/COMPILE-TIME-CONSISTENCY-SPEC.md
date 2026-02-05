# Compile-Time Consistency Specification

## Implementation Status

> **✅ COMPLETED** - Phase G1 Code Generation (February 2026)

### Completed

- ✅ **Phase G0**: `[ServiceCategory]` and `[ServiceAction]` attributes added to Core
- ✅ **Generator Infrastructure**: Three generator projects created
  - `ExcelMcp.Generators` - Core generator (produces ServiceRegistry)
  - `ExcelMcp.Generators.Shared` - Shared models and extractors
  - `ExcelMcp.Generators.Cli` - CLI generator (not producing output - architectural limitation)
  - `ExcelMcp.Generators.Mcp` - MCP generator (not producing output - architectural limitation)
- ✅ **PowerQuery Fully Converted**:
  - `IPowerQueryCommands.cs` has `[ServiceCategory]` and `[ServiceAction]` attributes
  - Generator produces: enum, constants, CliSettings, RouteCliArgs, RouteAction, ValidActions, ForwardMethods
  - `PowerQueryCommand.cs` uses generated `ServiceRegistry.PowerQuery.CliSettings` and `ServiceCommandBase<T>`
  - `ExcelPowerQueryTool.cs` uses generated `RouteAction` method
- ✅ **ServiceCommandBase<T>**: Shared CLI base class for session/action validation
- ✅ **McpMeta Attributes**: Kept as-is (exposed to MCP clients for tool categorization)
- ✅ **All Interfaces Converted**: All command interfaces have `[ServiceCategory]` and `[ServiceAction]` attributes

### Remaining Interfaces to Convert

| Interface | CLI Command | MCP Tool | Status |
|-----------|-------------|----------|--------|
| `IPowerQueryCommands` | `PowerQueryCommand` | `ExcelPowerQueryTool` | ✅ Done |
| `ICalculationModeCommands` | `CalculationModeCommand` | `ExcelCalculationModeTool` | ✅ Done |
| `ISheetCommands` + `ISheetStyleCommands` | `SheetCommand` | `ExcelWorksheetTool` + `ExcelWorksheetStyleTool` | ✅ Done |
| `IChartCommands` + `IChartConfigCommands` | `ChartCommand` + `ChartConfigCommand` | `ExcelChartTool` + `ExcelChartConfigTool` | ✅ Done |
| `IConnectionCommands` | `ConnectionCommand` | `ExcelConnectionTool` | ✅ Done |
| `INamedRangeCommands` | `NamedRangeCommand` | `ExcelNamedRangeTool` | ✅ Done |
| `IConditionalFormattingCommands` | `ConditionalFormatCommand` | `ExcelConditionalFormatTool` | ✅ Done |
| `IVbaCommands` | `VbaCommand` | `ExcelVbaTool` | ✅ Done |
| `IDataModelCommands` | `DataModelCommand` | `ExcelDataModelTool` | ✅ Done |
| `IDataModelRelCommands` | `DataModelRelCommand` | `ExcelDataModelRelTool` | ✅ Done |
| `IRangeCommands` | `RangeCommand` | `ExcelRangeTool` | ✅ Done |
| `ITableCommands` | `TableCommand` | `ExcelTableTool` | ✅ Done |
| `IPivotTableCommands` + `IPivotTableFieldCommands` + `IPivotTableCalcCommands` | `PivotTableCommand` | `ExcelPivotTableTool` + `ExcelPivotTableFieldTool` + `ExcelPivotTableCalcTool` | ✅ Done |
| `ISlicerCommands` | `SlicerCommand` | `ExcelSlicerTool` | ✅ Done |

---

## Generator Output (Implemented)

The Core generator (`ExcelMcp.Generators`) produces the following for each `[ServiceCategory]` interface:

### Generated ServiceRegistry.{Category}.g.cs

```csharp
public static partial class ServiceRegistry
{
    public static partial class PowerQuery
    {
        // 1. Enum (all actions)
        public enum Action { List, View, Create, Update, Delete, Refresh, ... }
        
        // 2. Constants (category/action strings)
        public const string Category = "powerquery";
        
        public static class Actions
        {
            public const string List = "list";
            public const string View = "view";
            // ...
        }
        
        // 3. ValidActions (for CLI help text)
        public static readonly IReadOnlyList<string> ValidActions = [Actions.List, Actions.View, ...];
        
        // 4. ToActionString (enum → string)
        public static string ToActionString(Action action) => action switch
        {
            Action.List => Actions.List,
            // ...
        };
        
        // 5. CliSettings (Spectre.Console.Cli settings class)
#if SPECTRE_CONSOLE
        public class CliSettings : CommandSettings
        {
            [CommandOption("-s|--session <SESSION>")]
            public string? SessionId { get; init; }
            
            [CommandOption("-a|--action <ACTION>")]
            public string? Action { get; init; }
            
            [CommandOption("-n|--query-name <NAME>")]
            public string? QueryName { get; init; }
            // ... all parameters from interface methods
        }
#endif
        
        // 6. RouteCliArgs (CliSettings → service command + args)
        public static (string command, object? args) RouteCliArgs(
            CliSettings settings,
            string action,
            Func<string?, string?> resolveFileOrValue)
        {
            return action switch
            {
                Actions.List => ($"{Category}.{Actions.List}", new { }),
                Actions.View => ($"{Category}.{Actions.View}", new { queryName = settings.QueryName }),
                // ...
            };
        }
        
        // 7. RouteAction (MCP action → ForwardToService call)
        public static string RouteAction(
            Action action,
            string sessionId,
            Func<string, string, object?, string> forward,
            /* all parameters */)
        {
            return action switch
            {
                Action.List => forward($"{Category}.{Actions.List}", sessionId, null),
                Action.View => forward($"{Category}.{Actions.View}", sessionId, new { queryName }),
                // ...
            };
        }
        
        // 8. ForwardMethods (individual strongly-typed methods)
        public static string ForwardList(string sessionId, Func<...> forward) => ...;
        public static string ForwardView(string sessionId, string queryName, Func<...> forward) => ...;
    }
}
```

### What Cannot Be Generated (Architectural Limitations)

1. **Full CliCommand class** - Would require referencing Service project from Core (circular dependency)
2. **CLI/MCP generator output** - Generators only see their own project's sources, not Core
3. **McpMeta attributes** - These are MCP SDK metadata for client-side filtering, kept manual

---

## McpMeta Attributes Decision

**Decision: KEEP** (February 2026)

The `[McpMeta]` attributes on MCP tools are **not consumed by our code** but are exposed to MCP clients:

```csharp
[McpMeta("category", "query")]        // Tool categorization
[McpMeta("requiresSession", true)]    // Session requirement hint
[McpMeta("fileFormat", ".xlsm")]      // VBA tool only
```

**Categories in use:** query, data, analysis, session, structure, settings, automation

**Purpose:**
- Exposed in MCP tool schema JSON
- MCP clients can filter/categorize tools
- LLMs may use metadata for tool selection
- Documents tool intent

**Not generated** because:
- Static metadata, not derived from interface
- Would require additional attributes on interfaces with unclear benefit

---

## Problem Statement

The MCP Server, CLI, and Service all expose the same Core functionality but use different mechanisms:
- **MCP Server**: 20+ tool classes with enum-based action routing
- **CLI**: 15+ command classes registered in Program.cs
- **Service**: String-based category/action routing via switch statements

This creates fragility at **three levels**:

### Level 1: Tools/Categories
| MCP Server | CLI | Service |
|------------|-----|---------|
| `ExcelPowerQueryTool.cs` | `PowerQueryCommand` in Program.cs | `"powerquery"` case in category switch |

**Risk:** Add a new MCP tool but forget to add the Service handler = runtime failure

### Level 2: Actions within Tools
| MCP Server | CLI | Service |
|------------|-----|---------|
| `PowerQueryAction.Evaluate` enum | `--action evaluate` | `"evaluate"` case in action switch |

**Risk:** Typo in `ForwardToService("powerquery.evalute")` compiles but fails at runtime

### Level 3: Parameters
| MCP Server | CLI | Service |
|------------|-----|---------|
| `string? queryName` parameter | `--query-name` option | `args.QueryName` deserialization |

**Risk:** Rename parameter in MCP but forget to update Service = null values at runtime

## Proposed Solution: Strongly-Typed Service Protocol

### Single Source of Truth: ServiceRegistry

Create a central registry that defines **all** tools, actions, and parameter contracts:

```csharp
// ServiceRegistry.cs - THE source of truth for the entire system
public static class ServiceRegistry
{
    // === TOOL CATEGORIES ===
    public static class Categories
    {
        public const string Service = "service";
        public const string Session = "session";
        public const string Sheet = "sheet";
        public const string Range = "range";
        public const string Table = "table";
        public const string PowerQuery = "powerquery";
        public const string PivotTable = "pivottable";
        public const string Chart = "chart";
        public const string ChartConfig = "chartconfig";
        public const string Connection = "connection";
        public const string Calculation = "calculation";
        public const string NamedRange = "namedrange";
        public const string ConditionalFormat = "conditionalformat";
        public const string Vba = "vba";
        public const string DataModel = "datamodel";
        public const string DataModelRel = "datamodelrel";
        public const string Slicer = "slicer";
        
        public static readonly string[] All = 
        [
            Service, Session, Sheet, Range, Table, PowerQuery, PivotTable,
            Chart, ChartConfig, Connection, Calculation, NamedRange,
            ConditionalFormat, Vba, DataModel, DataModelRel, Slicer
        ];
    }
    
    // === OPERATIONS (Category + Action) ===
    public static class PowerQuery
    {
        public static ServiceOperation List => new(Categories.PowerQuery, "list");
        public static ServiceOperation View => new(Categories.PowerQuery, "view");
        public static ServiceOperation Create => new(Categories.PowerQuery, "create");
        public static ServiceOperation Update => new(Categories.PowerQuery, "update");
        public static ServiceOperation Delete => new(Categories.PowerQuery, "delete");
        public static ServiceOperation Refresh => new(Categories.PowerQuery, "refresh");
        public static ServiceOperation RefreshAll => new(Categories.PowerQuery, "refresh-all");
        public static ServiceOperation Evaluate => new(Categories.PowerQuery, "evaluate");
        public static ServiceOperation LoadTo => new(Categories.PowerQuery, "load-to");
        public static ServiceOperation Unload => new(Categories.PowerQuery, "unload");
        public static ServiceOperation Rename => new(Categories.PowerQuery, "rename");
        public static ServiceOperation GetLoadConfig => new(Categories.PowerQuery, "get-load-config");
        
        public static readonly ServiceOperation[] All = 
        [
            List, View, Create, Update, Delete, Refresh, RefreshAll,
            Evaluate, LoadTo, Unload, Rename, GetLoadConfig
        ];
    }
    
    // ... similar for all categories
}
```

### Phase 1: Tool/Category Consistency

**Goal:** Compile-time verification that all categories exist in Service

**Implementation:**

1. Service category switch uses `Categories.All`:

```csharp
// ExcelMcpService.cs - Generated or validated against ServiceRegistry
return category switch
{
    ServiceRegistry.Categories.Service => HandleServiceCommand(action),
    ServiceRegistry.Categories.Session => HandleSessionCommand(action, request),
    ServiceRegistry.Categories.Sheet => await HandleSheetCommandAsync(action, request),
    // ... all from ServiceRegistry.Categories
    _ => throw new InvalidOperationException($"Unknown category: {category}. Valid: {string.Join(", ", ServiceRegistry.Categories.All)}")
};
```

2. Unit test validates completeness:

```csharp
[Fact]
public void AllCategoriesHaveHandlers()
{
    var service = new ExcelMcpService();
    foreach (var category in ServiceRegistry.Categories.All)
    {
        // Send ping to each category, verify it doesn't return "Unknown category"
        var response = service.HandleRequest($"{category}.ping");
        Assert.DoesNotContain("Unknown category", response.ErrorMessage ?? "");
    }
}
```

### Phase 2: Action Consistency (Updated)

**Goal:** Compile-time verification that all actions per category exist in Service

**Change MCP ForwardToService calls:**

```csharp
// Before (brittle)
return ExcelToolsBase.ForwardToService("powerquery.list", sessionId);

// After (type-safe)
return ExcelToolsBase.ForwardToService(ServiceRegistry.PowerQuery.List, sessionId);
```

**Service validates at startup:**

```csharp
// ExcelMcpService constructor
static ExcelMcpService()
{
    // Validate all registered operations have handlers
    ValidateAllOperationsHaveHandlers();
}

private static void ValidateAllOperationsHaveHandlers()
{
    var allOperations = typeof(ServiceRegistry)
        .GetNestedTypes()
        .Where(t => t.Name != "Categories")
        .SelectMany(t => t.GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(f => f.FieldType == typeof(ServiceOperation))
            .Select(f => (ServiceOperation)f.GetValue(null)!));
    
    // This runs at class load time - fails fast if any operation is missing
    foreach (var op in allOperations)
    {
        if (!_handlers.ContainsKey(op.ToProtocolString()))
        {
            throw new InvalidOperationException(
                $"ServiceRegistry defines {op.ToProtocolString()} but no handler exists!");
        }
    }
}
```

### Phase 3: Parameter Consistency

**Goal:** Type-safe parameter passing

**Define request DTOs in ServiceRegistry:**

```csharp
public static class ServiceRegistry
{
    public static class PowerQuery
    {
        // Operations (as before)
        public static ServiceOperation Create => new(Categories.PowerQuery, "create");
        
        // Request DTOs - shared between MCP and Service
        public record CreateRequest(
            string QueryName,
            string MCode,
            string? LoadDestination = null,
            string? TargetSheet = null
        );
    }
}
```

**MCP Tool uses DTO:**

```csharp
private static string ForwardCreate(string sessionId, string queryName, string mCode, string? loadDestination)
{
    var request = new ServiceRegistry.PowerQuery.CreateRequest(queryName, mCode, loadDestination);
    return ExcelToolsBase.ForwardToService(ServiceRegistry.PowerQuery.Create, sessionId, request);
}
```

**Service deserializes same DTO:**

```csharp
private async Task<ServiceResponse> HandlePowerQueryCreate(ServiceRequest request)
{
    var args = request.DeserializeArgs<ServiceRegistry.PowerQuery.CreateRequest>();
    // args.QueryName, args.MCode, etc. - all type-safe
}
```

## Alternative Approaches

### Option B: Roslyn Analyzer (Less Invasive)

Keep current string-based approach but add a custom Roslyn analyzer:

```csharp
[assembly: ExcelMcpAnalyzer]

// Analyzer validates:
// 1. All ForwardToService() calls use valid operation strings
// 2. All Service switch cases match known operations
// 3. All CLI commands map to known operations
```

**Pros:** No code changes, just add analyzer
**Cons:** More complex to maintain, analyzer development overhead

### Option C: Exhaustive Switch on Enums in Service

Change Service to use exhaustive switch expressions on enums:

```csharp
// Service receives enum instead of string
private async Task<ServiceResponse> HandlePowerQueryCommandAsync(PowerQueryAction action, ServiceRequest request)
{
    return action switch
    {
        PowerQueryAction.List => ...,
        PowerQueryAction.View => ...,
        // Compiler error if any enum value is missing!
    };
}
```

**Requires:** Protocol change to send enum values instead of strings

## Option D: Code Generation from Core (Preferred)

**Insight:** Core is the actual source of truth - it has the real `*Commands` classes with the actual implementations. CLI and MCP are just interfaces to Core.

### Data Sources

| Layer | Source of Truth For | How to Extract |
|-------|---------------------|----------------|
| **Core** | Operations that exist | Reflect on `I*Commands` interfaces and `*Commands` classes |
| **Core** | Parameter types | Method signatures on command interfaces |
| **MCP** | Tool descriptions | XML docs on `[McpServerTool]` methods |
| **CLI** | Help text | `[Description]` attributes on Spectre.Console commands |

### Generation Flow

```
┌─────────────────────────────────────────────────────────────┐
│  SOURCE OF TRUTH                                            │
├─────────────────────────────────────────────────────────────┤
│  Core I*Commands Interfaces                                 │
│  - IPowerQueryCommands.ListAsync()                         │
│  - IPowerQueryCommands.CreateAsync(name, mCode, ...)       │
│  - IRangeCommands.GetValuesAsync(range, sheet)             │
└───────────────┬─────────────────────────────────────────────┘
                │
                ▼ Source Generator / T4 Template
┌─────────────────────────────────────────────────────────────┐
│  GENERATED                                                  │
├─────────────────────────────────────────────────────────────┤
│  ServiceRegistry.cs                                         │
│  - Categories, Operations, Request DTOs                     │
│                                                             │
│  ServiceHandlers.generated.cs                               │
│  - Switch cases for all operations                          │
│  - Calls to Core commands                                   │
└─────────────────────────────────────────────────────────────┘
```

### Metadata Attributes on Core

Add attributes to Core interfaces to capture metadata:

```csharp
// New attributes in ExcelMcp.Core
[AttributeUsage(AttributeTargets.Interface)]
public class ServiceCategoryAttribute(string category) : Attribute
{
    public string Category { get; } = category;
}

[AttributeUsage(AttributeTargets.Method)]
public class ServiceActionAttribute(string action) : Attribute
{
    public string Action { get; } = action;
}

// Applied to Core interfaces
[ServiceCategory("powerquery")]
public interface IPowerQueryCommands
{
    [ServiceAction("list")]
    Task<PowerQueryListResult> ListAsync(IExcelBatch batch);
    
    [ServiceAction("create")]
    Task<OperationResult> CreateAsync(
        IExcelBatch batch,
        string queryName,
        string mCode,
        string? loadDestination = null);
    
    [ServiceAction("view")]
    Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName);
}
```

### Source Generator Output

The generator reads Core interfaces and produces:

```csharp
// Generated: ServiceRegistry.generated.cs
public static partial class ServiceRegistry
{
    public static class Categories
    {
        public const string PowerQuery = "powerquery";
        public const string Range = "range";
        // ... generated from [ServiceCategory] attributes
    }
    
    public static class PowerQuery
    {
        public static ServiceOperation List => new(Categories.PowerQuery, "list");
        public static ServiceOperation Create => new(Categories.PowerQuery, "create");
        public static ServiceOperation View => new(Categories.PowerQuery, "view");
        
        // Generated from method signature
        public record CreateRequest(
            string QueryName,
            string MCode,
            string? LoadDestination = null
        );
    }
}

// Generated: ServiceHandlers.generated.cs
public partial class ExcelMcpService
{
    private async Task<ServiceResponse> HandlePowerQueryAsync(string action, ServiceRequest request)
    {
        return action switch
        {
            "list" => await HandlePowerQueryListAsync(request),
            "create" => await HandlePowerQueryCreateAsync(request),
            "view" => await HandlePowerQueryViewAsync(request),
            _ => ServiceResponse.Error($"Unknown powerquery action: {action}")
        };
    }
    
    private async Task<ServiceResponse> HandlePowerQueryCreateAsync(ServiceRequest request)
    {
        var args = request.DeserializeArgs<ServiceRegistry.PowerQuery.CreateRequest>();
        var result = await _powerQueryCommands.CreateAsync(
            request.Batch,
            args.QueryName,
            args.MCode,
            args.LoadDestination);
        return ServiceResponse.FromResult(result);
    }
}
```

### Documentation Extraction

For tool descriptions, create a separate analyzer that validates MCP/CLI match Core:

```csharp
// Build-time validation (not code generation)
// Runs as unit test or MSBuild task

[Fact]
public void AllCoreOperationsHaveMcpTools()
{
    var coreOperations = ExtractFromCore(); // Uses reflection on [ServiceAction]
    var mcpOperations = ExtractFromMcp();   // Uses reflection on [McpServerTool]
    
    foreach (var op in coreOperations)
    {
        Assert.Contains(mcpOperations, m => 
            m.Category == op.Category && m.Action == op.Action);
    }
}

[Fact]
public void AllCoreOperationsHaveCliCommands()
{
    var coreOperations = ExtractFromCore();
    var cliCommands = ExtractFromCli();     // Uses reflection on Spectre commands
    
    // Validate CLI covers all Core operations
}
```

### Benefits Over ServiceRegistry Approach

| Aspect | Manual ServiceRegistry | Generated from Core |
|--------|------------------------|---------------------|
| Source of truth | New file to maintain | Core interfaces (already exist) |
| Adding operation | Update 3+ files | Add method to Core, regenerate |
| Parameter changes | Update DTOs manually | Regenerated from method signature |
| Forgotten updates | Runtime error | Compile error (missing handler) |
| Documentation | Separate maintenance | Extracted from existing docs |

### Implementation Phases

| Phase | Description | Effort |
|-------|-------------|--------|
| **G0** | Add `[ServiceCategory]` and `[ServiceAction]` attributes to Core | Low |
| **G1** | Create source generator for `ServiceRegistry.generated.cs` | Medium |
| **G2** | Generate Service handlers | Medium |
| **G3** | Generate MCP tool routing (optional - may keep manual) | High |
| **G4** | Validation tests for CLI/MCP coverage | Low |

### Files to Modify

**Phase G0 (Attributes):**
- `src/ExcelMcp.Core/Attributes/ServiceCategoryAttribute.cs` (new)
- `src/ExcelMcp.Core/Attributes/ServiceActionAttribute.cs` (new)
- `src/ExcelMcp.Core/Commands/I*Commands.cs` (add attributes)

**Phase G1 (Generator):**
- `generators/ExcelMcp.Generators/ServiceRegistryGenerator.cs` (new project)
- `src/ExcelMcp.ComInterop/ServiceClient/ServiceRegistry.generated.cs` (generated)

**Phase G2 (Handlers):**
- `src/ExcelMcp.Service/ExcelMcpService.generated.cs` (generated)

**Phase G4 (Validation):**
- `tests/ExcelMcp.Consistency.Tests/CoreMcpConsistencyTests.cs` (new)
- `tests/ExcelMcp.Consistency.Tests/CoreCliConsistencyTests.cs` (new)

## Recommendation

**Option D (Code Generation from Core)** is the cleanest long-term solution because:

1. **Core is already the source of truth** - we're just making it explicit
2. **No new files to maintain** - generator reads existing interfaces
3. **Parameter types flow automatically** - method signatures become DTOs
4. **Impossible to forget** - adding a Core method triggers regeneration

**Suggested first step:** Add `[ServiceCategory]` and `[ServiceAction]` attributes to Core interfaces. This is low-risk and immediately documents the mapping, even before building the generator.

## Implementation Priority

| Phase | Level | Change | Effort | Impact |
|-------|-------|--------|--------|--------|
| **P0** | Categories | Create `ServiceRegistry.Categories` constants | Low | Typos in category strings → compile error |
| **P1** | Actions | Create `ServiceRegistry.{Category}.{Action}` operations | Medium | Typos in action strings → compile error |
| **P2** | Actions | Validate all operations have handlers at startup | Low | Missing handler → fail-fast at startup |
| **P3** | Parameters | Create shared request DTOs | Medium | Parameter mismatches → compile error |
| **P4** | All | CLI command validation via reflection | Medium | CLI drift → test failure |
| **P5** | All | Roslyn analyzer for cross-project validation | High | All inconsistencies → compile error |

## Files to Modify

**Phase 0 (Categories):**
- `src/ExcelMcp.ComInterop/ServiceClient/ServiceRegistry.cs` (new - Categories only)
- `src/ExcelMcp.Service/ExcelMcpService.cs` (use ServiceRegistry.Categories constants)

**Phase 1 (Actions):**
- `src/ExcelMcp.ComInterop/ServiceClient/ServiceRegistry.cs` (add operation definitions)
- `src/ExcelMcp.McpServer/Tools/*.cs` (update ForwardToService calls)
- `src/ExcelMcp.McpServer/Tools/ExcelToolsBase.cs` (update signature)

**Phase 2 (Validation):**
- `src/ExcelMcp.Service/ExcelMcpService.cs` (add startup validation)
- `tests/ExcelMcp.Service.Tests/ServiceRegistryValidationTests.cs` (new)

**Phase 3 (Parameters):**
- `src/ExcelMcp.ComInterop/ServiceClient/ServiceRegistry.cs` (add request DTOs)
- All MCP tools and Service handlers (use DTOs)

**Phase 4 (CLI):**
- `tests/ExcelMcp.CLI.Tests/CliServiceConsistencyTests.cs` (new)

**Phase 5 (Analyzer):**
- `analyzers/ExcelMcp.Analyzers/` (new project)

## Example: Phase 1 Implementation

```csharp
// Before (brittle)
private static string ForwardList(string sessionId)
{
    return ExcelToolsBase.ForwardToService("powerquery.list", sessionId);
}

// After (compile-time safe)
private static string ForwardList(string sessionId)
{
    return ExcelToolsBase.ForwardToService(ServiceAction.PowerQuery.List, sessionId);
}
```

If you typo `ServiceAction.PowerQuery.Listt`, the compiler catches it immediately.

## Testing Strategy

After implementing Phase 1:
1. Delete `check-service-consistency.ps1` pre-commit check
2. Add unit test that validates all `ServiceAction` values have corresponding Service handlers
3. The test uses reflection to enumerate `ServiceAction` and verify each has a handler

```csharp
[Fact]
public void AllServiceActionsHaveHandlers()
{
    var allActions = typeof(ServiceAction)
        .GetNestedTypes()
        .SelectMany(t => t.GetProperties(BindingFlags.Public | BindingFlags.Static))
        .Select(p => ((ServiceOperation)p.GetValue(null)!).ToProtocolString());
    
    foreach (var action in allActions)
    {
        // Send to service, verify it doesn't return "Unknown action"
    }
}
```

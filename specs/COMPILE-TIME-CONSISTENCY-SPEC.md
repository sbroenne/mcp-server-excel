# Compile-Time Consistency: What Was Implemented

> **Status**: ✅ COMPLETED (February 2026)

## Overview

This document describes the **code generation system** that ensures consistency between Core interfaces, CLI commands, and MCP tools.

**Single Source of Truth**: Core interface methods annotated with `[ServiceCategory]` and `[ServiceAction]` attributes.

**Generated Code**: A Roslyn source generator produces `ServiceRegistry.{Category}.g.cs` files with:
- Action enums
- String constants
- CLI Settings classes
- Routing methods

---

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                          CORE (Source of Truth)                             │
│                                                                             │
│   [ServiceCategory("powerquery")]                                          │
│   public interface IPowerQueryCommands                                     │
│   {                                                                         │
│       [ServiceAction("list")]                                              │
│       List<QueryInfo> List(IExcelBatch batch);                             │
│                                                                             │
│       [ServiceAction("create")]                                            │
│       void Create(IExcelBatch batch, string queryName, string mCode, ...); │
│   }                                                                         │
└──────────────────────────────────┬──────────────────────────────────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │    SOURCE GENERATOR         │
                    │    (ExcelMcp.Generators)    │
                    └──────────────┬──────────────┘
                                   │
         ┌─────────────────────────┼─────────────────────────┐
         │                         │                         │
         ▼                         ▼                         ▼
┌─────────────────┐   ┌─────────────────────┐   ┌─────────────────────────┐
│  ENUM GENERATED │   │  CLI USES GENERATED │   │  MCP USES GENERATED     │
│                 │   │                     │   │                         │
│  PowerQueryAction│   │  CliSettings class  │   │  RouteAction() method   │
│  - List          │   │  RouteCliArgs()     │   │  Forward{Action}()      │
│  - Create        │   │  ValidActions[]     │   │                         │
│  - Refresh       │   │                     │   │                         │
└─────────────────┘   └─────────────────────┘   └─────────────────────────┘
         │                         │                         │
         └─────────────────────────┼─────────────────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │    ALL BUILD THE SAME       │
                    │    (command, args) TUPLE    │
                    └──────────────┬──────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │    SERVICE (ExcelMcpService)│
                    │    Routes to Core Commands  │
                    └─────────────────────────────┘
```

---

## What Gets Generated

For each `[ServiceCategory]` interface, the generator produces `ServiceRegistry.{Category}.g.cs`:

### 1. Action Enum
```csharp
public enum PowerQueryAction
{
    List,
    View,
    Create,
    Update,
    Delete,
    Refresh,
    // ... one per [ServiceAction] method
}
```

### 2. String Constants
```csharp
public const string Category = "powerquery";
public const string McpToolName = "excel_powerquery";

// Action constants
public const string ListAction = "list";
public const string CreateAction = "create";

// Full command strings  
public const string ListCommand = "powerquery.list";
public const string CreateCommand = "powerquery.create";

public static readonly string[] ValidActions = ["list", "view", "create", ...];
```

### 3. Parsing/Conversion Methods
```csharp
// String → Enum
public static bool TryParseAction(string actionString, out PowerQueryAction action)

// Enum → String  
public static string ToActionString(PowerQueryAction action)
```

### 4. CLI Settings Class
```csharp
public sealed class CliSettings : Spectre.Console.Cli.CommandSettings
{
    [CommandArgument(0, "<ACTION>")]
    public string Action { get; init; } = string.Empty;

    [CommandOption("-s|--session <SESSION>")]
    public string SessionId { get; init; } = string.Empty;

    [CommandOption("--query-name <QUERYNAME>")]
    public string? QueryName { get; init; }

    [CommandOption("--m-code <MCODE>")]
    public string? MCode { get; init; }
    
    // ... all parameters from all methods in the interface
}
```

### 5. CLI Routing Method
```csharp
public static (string Command, object? Args) RouteCliArgs(
    string action,
    string? queryName = null,
    string? mCode = null,
    // ... all parameters
)
{
    var command = $"powerquery.{action}";
    object? args = action switch
    {
        "list" => null,
        "create" => new { queryName, mCode, loadDestination, ... },
        // ... one case per action
    };
    return (command, args);
}
```

### 6. MCP Routing Method
```csharp
public static string RouteAction(
    PowerQueryAction action,
    string sessionId,
    Func<string, string, object?, string> forwardToService,
    // ... all parameters
)
{
    return action switch
    {
        PowerQueryAction.List => ForwardList(sessionId, forwardToService),
        PowerQueryAction.Create => ForwardCreate(sessionId, forwardToService, queryName, mCode, ...),
        // ...
    };
}
```

### 7. Forward Methods (one per action)
```csharp
public static string ForwardList(string sessionId, Func<...> forwardToService)
{
    return forwardToService("powerquery.list", sessionId, null);
}

public static string ForwardCreate(
    string sessionId,
    Func<...> forwardToService,
    string? queryName = null,
    string? mCode = null,
    // ...
)
{
    ParameterTransforms.RequireNotEmpty(queryName, "queryName", "create");
    return forwardToService("powerquery.create", sessionId, new
    {
        QueryName = queryName,
        MCode = mCode,
        // ...
    });
}
```

---

## How CLI Uses Generated Code

**Before (manual)**:
```csharp
// Each CLI command had to manually build args objects matching Service expectations
```

**After (generated)**:
```csharp
internal sealed class PowerQueryCommand : ServiceCommandBase<ServiceRegistry.PowerQuery.CliSettings>
{
    protected override string? GetSessionId(CliSettings s) => s.SessionId;
    protected override string? GetAction(CliSettings s) => s.Action;
    protected override IReadOnlyList<string> ValidActions => ServiceRegistry.PowerQuery.ValidActions;

    protected override (string command, object? args) Route(CliSettings settings, string action)
    {
        return ServiceRegistry.PowerQuery.RouteCliArgs(
            action,
            queryName: settings.QueryName,
            mCode: settings.MCode,
            // ... settings properties map 1:1 to generated parameters
        );
    }
}
```

---

## How MCP Uses Generated Code

**Before (manual)**:
```csharp
// Each MCP tool had switch statements calling ForwardToService with string literals
return ExcelToolsBase.ForwardToService("powerquery.create", sessionId, new { queryName, mCode });
```

**After (generated)**:
```csharp
public static partial string ExcelPowerQuery(PowerQueryAction action, string sessionId, ...)
{
    return ExcelToolsBase.ExecuteToolAction(
        "excel_powerquery",
        ServiceRegistry.PowerQuery.ToActionString(action),
        () => ServiceRegistry.PowerQuery.RouteAction(
            action,
            sessionId,
            ExcelToolsBase.ForwardToServiceFunc,
            queryName: queryName,
            mCode: mCode,
            // ... MCP parameters map 1:1 to generated parameters
        ));
}
```

---

## Compile-Time Guarantees

| What | How It's Enforced |
|------|-------------------|
| Action enum values exist for all Core methods | Generator reads `[ServiceAction]` attributes |
| CLI settings have properties for all parameters | Generator extracts method parameters |
| RouteCliArgs covers all actions | Generator creates switch case per method |
| RouteAction covers all enum values | Generator creates switch case per enum |
| Parameter names match across layers | Same generator produces all routing code |
| String constants are consistent | Generated once from source of truth |

---

## What's NOT Generated (By Design)

1. **Service routing logic** - `ExcelMcpService.cs` still has manual switch statements (routes to Core)
2. **MCP tool class definition** - The `[McpServerTool]` attribute and parameter list are manual
3. **McpMeta attributes** - Static metadata for MCP clients, not derived from interface

---

## Generator Projects

| Project | Purpose |
|---------|---------|
| `ExcelMcp.Generators` | Main generator - produces ServiceRegistry files |
| `ExcelMcp.Generators.Shared` | Shared models (ServiceInfo, MethodInfo, ParameterInfo) |
| `ExcelMcp.Generators.Mcp` | (Placeholder - architectural limitation prevents MCP-specific generation) |
| `ExcelMcp.Generators.Cli` | (Placeholder - architectural limitation prevents CLI-specific generation) |

**Limitation**: Roslyn source generators can only see the project they're attached to. We attach to Core, so we can only generate into Core (the ServiceRegistry namespace).

---

## The Remaining Gap: Service Routing

The Service (`ExcelMcpService.cs`) still has **manual** switch statements that route commands to Core:

```csharp
// This is NOT generated - must be manually maintained
private Task<ServiceResponse> HandlePowerQueryCommandAsync(string action, ServiceRequest request)
{
    if (ServiceRegistry.PowerQuery.TryParseAction(action, out var pqAction))
    {
        return WithSessionAsync(request.SessionId, batch => pqAction switch
        {
            PowerQueryAction.List => SerializeResult(_powerQueryCommands.List(batch)),
            PowerQueryAction.Create => ExecuteCreate(batch, request),
            // ... every action manually mapped
        });
    }
    return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {action}" });
}
```

**Why not generated?**
- Would require the generator to understand Core method signatures deeply
- Return type handling varies (some return data, some are void)
- Error handling patterns differ per method

**Testing strategy for this gap**: The existing pre-commit scripts verify:
- All enum actions have Service implementations (`check-mcp-core-implementations.ps1`)
- All CLI actions have handlers (`check-cli-action-coverage.ps1`)

---

## Future Consideration: Service Generation

If we wanted to generate Service routing, we'd need:

1. **Standardized Core method signatures** - All methods return `OperationResult<T>` or `void`
2. **Attribute-based parameter binding** - Mark which JSON property maps to which parameter
3. **Separate generator attached to Service project** - Can see both Core and Service

This is **not implemented** because the current manual approach is:
- Well-tested via pre-commit checks
- Clear and debuggable
- Only ~15 lines per action (acceptable maintenance cost)


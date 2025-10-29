# Layer Separation Analysis - Core, MCP Server, CLI

> **Date**: 2025-01-29  
> **Purpose**: Verify proper separation of concerns across layers  
> **Finding**: VIOLATIONS DETECTED - Mixed responsibilities

## Current State Analysis

### ✅ What's Working

1. **MCP Server Layer** - Properly isolated
   - No CLI-specific code found (`excelcli`, `pq-list`, etc.)
   - Uses MCP-specific attributes (`[McpServerTool]`, `[Required]`, etc.)
   - Only references Core for business logic

2. **CLI Layer** - Properly isolated
   - No MCP-specific code
   - Uses Spectre.Console for presentation
   - Only references Core for business logic

### ❌ Violations Found

#### 1. Core Contains Layer-Specific Suggestions

**Location**: `src/ExcelMcp.Core/Models/ResultTypes.cs`
```csharp
public abstract class ResultBase
{
    /// <summary>
    /// Suggested next actions for LLM workflow guidance  // ← MCP-SPECIFIC COMMENT
    /// </summary>
    public List<string> SuggestedNextActions { get; set; } = new();
    
    /// <summary>
    /// Contextual workflow hint for LLM  // ← MCP-SPECIFIC COMMENT
    /// </summary>
    public string? WorkflowHint { get; set; }
}
```

**Problem**: Comments explicitly mention "LLM workflow guidance" but property is populated differently by Core, MCP, and CLI.

**Impact**: Core layer is aware of presentation layer concerns.

#### 2. Core Contains CLI Command Names

**Location**: `src/ExcelMcp.Core/Commands/ParameterCommands.cs`
```csharp
result.SuggestedNextActions = new List<string>
{
    "Use 'param-list' to see available parameters",  // ← CLI command name
    "Use 'param-create' to create a new named range"  // ← CLI command name
};
```

**Found in 19 Core files** with CLI-specific command names:
- `param-list`, `param-create`, `param-get`, `param-set`
- `pq-list`, `pq-view`, `pq-import`, `pq-update`, `pq-delete`
- `table-list`, `table-create`, etc.

**Problem**: Core knows about CLI presentation layer command naming.

#### 3. Core Contains MCP Action Names

**Location**: `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs` (and MCP Server Tools)
```csharp
// In Core:
result.SuggestedNextActions = new List<string>
{
    "Use 'view' to inspect a query's M code",  // ← MCP action name
    "Use 'import' to add a new Power Query"    // ← MCP action name
};

// In MCP Server:
result.SuggestedNextActions = new List<string>
{
    "Use 'view' to inspect a query's M code",  // ← Same suggestion!
    "Use 'import' to add a new Power Query"
};
```

**Found in**: 
- Core Commands: 19 files with MCP action names (`view`, `list`, `import`, etc.)
- MCP Server Tools: 6 files duplicating Core suggestions

**Problem**: Core and MCP Server both populate suggestions, causing duplication.

#### 4. Mixed Suggestion Styles in Core

**Location**: Various Core command files

```csharp
// Example 1: MCP-style (action names)
result.SuggestedNextActions = new List<string>
{
    "Use 'view' to inspect a query's M code"
};

// Example 2: CLI-style (command names)
result.SuggestedNextActions = new List<string>
{
    "Use 'param-list' to see available parameters"
};

// Example 3: Generic (layer-agnostic)
result.SuggestedNextActions = new List<string>
{
    "Check that the Excel file exists and is accessible"
};
```

**Problem**: Core doesn't know which layer will consume it, so suggestions are inconsistent.

## Architectural Issues

### Issue 1: Tight Coupling

```
┌─────────────────────────────────────────────┐
│              Core Layer                     │
│  ┌────────────────────────────────────┐    │
│  │ PowerQueryCommands.cs              │    │
│  │                                     │    │
│  │ result.SuggestedNextActions =      │    │
│  │   "Use 'pq-list' to see queries"   │ ←──┼── Knows CLI command!
│  │   "Use 'view' to inspect M code"   │ ←──┼── Knows MCP action!
│  └────────────────────────────────────┘    │
└─────────────────────────────────────────────┘
         ↓                    ↓
┌──────────────────┐  ┌──────────────────┐
│   CLI Layer      │  │  MCP Server      │
│                  │  │                  │
│ Displays         │  │ Overwrites or    │
│ Core suggestions │  │ adds to Core     │
│                  │  │ suggestions      │
└──────────────────┘  └──────────────────┘
```

**Impact**:
- Core must know about both CLI and MCP presentation
- Adding new CLI command requires Core changes
- Adding new MCP action requires Core changes
- Changing command/action names breaks suggestions in Core

### Issue 2: Responsibility Confusion

**Who populates SuggestedNextActions?**

1. **Core Commands** populate generic + CLI + MCP suggestions
2. **MCP Server Tools** sometimes add to Core suggestions
3. **CLI Commands** just display Core suggestions

**Result**: Unclear ownership, duplication, inconsistency.

### Issue 3: Validation Duplication

**Current state**:
```
┌─────────────────────────────────────────────────────┐
│ MCP Server                                           │
│ [RegularExpression("^(list|view|import|...)$")]     │ ← Hardcoded
│ [StringLength(255, MinimumLength = 1)]              │ ← Hardcoded
│ [FileExtensions(Extensions = "xlsx,xlsm")]          │ ← Hardcoded
└─────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────┐
│ CLI                                                  │
│ if (args.Length < 2) { return 1; }                  │ ← Manual
│ // No validation of query name, file extension, etc │
└─────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────┐
│ Core                                                 │
│ public Task<Result> ViewAsync(string queryName)     │ ← No validation
│ // Assumes inputs are valid                         │
└─────────────────────────────────────────────────────┘
```

**Problem**: Validation rules exist only in MCP Server, not shared.

## Recommended Architecture

### Proper Layer Separation

```
┌─────────────────────────────────────────────────────┐
│ Core Layer (Business Logic ONLY)                    │
│                                                      │
│ ✅ Excel COM operations                             │
│ ✅ Data transformations                             │
│ ✅ Business rules                                   │
│ ✅ Result objects (Success, ErrorMessage, Data)     │
│ ❌ NO presentation layer concerns                   │
│ ❌ NO command names (CLI or MCP)                    │
│ ❌ NO workflow guidance strings                     │
│                                                      │
│ public class ResultBase                             │
│ {                                                    │
│     public bool Success { get; set; }               │
│     public string? ErrorMessage { get; set; }       │
│     public string? FilePath { get; set; }           │
│     // NO SuggestedNextActions                      │
│     // NO WorkflowHint                              │
│ }                                                    │
└─────────────────────────────────────────────────────┘
         ↓                              ↓
┌─────────────────────┐    ┌───────────────────────────┐
│ CLI Layer           │    │ MCP Server Layer          │
│                     │    │                           │
│ ✅ Argument parsing │    │ ✅ MCP protocol handling  │
│ ✅ Console output   │    │ ✅ JSON serialization     │
│ ✅ CLI suggestions  │    │ ✅ MCP suggestions        │
│ ✅ Error formatting │    │ ✅ Tool metadata          │
│                     │    │                           │
│ Generates:          │    │ Generates:                │
│ "excelcli pq-list"  │    │ { tool: "excel_powerquery"│
│                     │    │   action: "list" }        │
└─────────────────────┘    └───────────────────────────┘
```

### Shared Validation (Proposed)

```
┌─────────────────────────────────────────────────────┐
│ Core.Validation (NEW - Shared)                      │
│                                                      │
│ public static class ValidationRules                 │
│ {                                                    │
│     public const int MaxQueryNameLength = 255;      │
│     public const string QueryNamePattern =          │
│         @"^[a-zA-Z0-9_]+$";                         │
│     public static string[] ValidActions =           │
│         { "list", "view", "import", ... };          │
│ }                                                    │
│                                                      │
│ public class ActionDefinition                       │
│ {                                                    │
│     public string Name { get; }                     │
│     public string CliCommand { get; }               │
│     public string McpAction { get; }                │
│     public ParameterDefinition[] Parameters { get; }│
│ }                                                    │
└─────────────────────────────────────────────────────┘
         ↓                              ↓
┌─────────────────────┐    ┌───────────────────────────┐
│ CLI                 │    │ MCP Server                │
│                     │    │                           │
│ Uses:               │    │ Uses:                     │
│ ValidationRules     │    │ ValidationRules           │
│ ActionDefinition    │    │ ActionDefinition          │
│   .CliCommand       │    │   .McpAction              │
└─────────────────────┘    └───────────────────────────┘
```

## Specific Violations by File

### Core Files with CLI Command Names

1. `src/ExcelMcp.Core/Commands/ParameterCommands.cs`
   - Lines with `param-list`, `param-create`, `param-get`, `param-set`

2. `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`
   - Lines with `pq-update`, `pq-export`, `pq-import`, `pq-delete`, `pq-loadto`, `pq-refresh`, `pq-test`

3. `src/ExcelMcp.Core/Commands/ConnectionCommands.cs`
   - Lines with `pq-export`, `pq-update`, `pq-delete`, `pq-loadto`

4. Multiple Table, DataModel, Script commands with similar violations

### Core Files with MCP Action Names

1. All Core Commands files populate `SuggestedNextActions` with action names like:
   - `"Use 'view' to inspect..."`
   - `"Use 'list' to see..."`
   - `"Use 'import' to add..."`

### MCP Server Files Duplicating Core Suggestions

1. `src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs`
2. `src/ExcelMcp.McpServer/Tools/ExcelParameterTool.cs`
3. `src/ExcelMcp.McpServer/Tools/TableTool.cs`
4. `src/ExcelMcp.McpServer/Tools/ExcelDataModelTool.cs`
5. `src/ExcelMcp.McpServer/Tools/ExcelVbaTool.cs`
6. `src/ExcelMcp.McpServer/Tools/ExcelConnectionTool.cs`

**Pattern**: MCP tools either:
- Overwrite Core suggestions entirely, OR
- Add to Core suggestions

**Inconsistency**: No clear pattern.

## Impact on Design Proposal

The new NextAction design must address these issues:

1. **Remove SuggestedNextActions from Core**
   - Core only returns business data
   - No workflow guidance in Core layer

2. **Move workflow guidance to presentation layers**
   - CLI generates CLI-specific suggestions
   - MCP generates MCP-specific suggestions
   - Core is unaware of presentation

3. **Share validation rules**
   - Single source of truth for parameter constraints
   - Used by both CLI and MCP Server
   - Core can optionally validate (business rules only)

4. **Share action definitions**
   - Central registry of actions
   - Maps to CLI commands and MCP actions
   - Single place to add/remove/rename actions

## Recommended Immediate Actions

1. **Accept current violations for v1.0**
   - Document as technical debt
   - Plan removal in v2.0

2. **New design must fix separation**
   - Core: No SuggestedNextActions, no WorkflowHint
   - Shared: Validation rules, action definitions
   - CLI: Generate CLI-specific suggestions
   - MCP: Generate MCP-specific suggestions

3. **Migration strategy**
   - Phase 1: Add new NextAction system alongside old
   - Phase 2: Migrate MCP/CLI to use new system
   - Phase 3: Remove SuggestedNextActions from Core (v2.0)

## Conclusion

**Finding**: ❌ **Separation of concerns is violated**

**Evidence**:
- Core contains 19 files with layer-specific suggestions
- CLI command names hardcoded in Core
- MCP action names hardcoded in Core
- Validation rules only in MCP Server, not shared

**Impact**: 
- High coupling between layers
- Difficult to maintain
- Duplication of effort
- Inconsistent suggestions

**Solution**: 
- Implement NextAction design with proper layer separation
- Remove presentation concerns from Core
- Share validation rules across layers
- Create central action definition registry

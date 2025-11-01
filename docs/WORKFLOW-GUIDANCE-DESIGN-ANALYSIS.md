# Workflow Guidance Design Analysis

## Current Design Issues

### Problem: Mixed Concerns Across Layers

**Core Layer** creates workflow guidance for **two completely different audiences**:
1. **CLI users** - Human terminal users who need command suggestions
2. **LLMs** - AI assistants that need semantic workflow hints

This violates separation of concerns and creates maintenance issues.

---

## Current Implementation

### Core Layer (`src/ExcelMcp.Core`)

**Files:**
- `Commands/PowerQueryWorkflowGuidance.cs` (269 lines)
- `Commands/DataModel/DataModelWorkflowGuidance.cs` (254 lines)
- `Commands/WorksheetWorkflowGuidance.cs` (minimal)

**Creates:**
- `SuggestedNextActions` - List of CLI command strings for humans
- `WorkflowHint` - Short description for humans/LLMs

**Examples:**
```csharp
// Power Query Import
result.SuggestedNextActions = [
    "Use 'set-load-to-table' with targetSheet parameter to load data to worksheet",
    "Or use 'set-load-to-data-model' to load to PowerPivot"
];
result.WorkflowHint = "Query imported as connection-only (M code not executed or validated).";

// Data Model Create Measure
result.SuggestedNextActions = [
    "Measure created successfully in Data Model",
    "Use 'dm-list-measures' to see all measures",
    "Use 'dm-view-measure' to inspect DAX formula"
];
result.WorkflowHint = "Measure created. Next, test the measure in a PivotTable or verify its formula.";
```

### CLI Layer (`src/ExcelMcp.CLI`)

**Usage:** Displays `SuggestedNextActions` directly to terminal users (116 occurrences)

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

### MCP Server Layer (`src/ExcelMcp.McpServer`)

**Usage:** **Overrides** Core's suggestions with LLM-specific guidance

**Problem:**
- Core creates CLI-focused suggestions like `"Use 'dm-list-measures' to see all measures"`
- MCP Server **ignores** these and creates its own LLM-specific suggestions
- Result: Core creates 269 lines of guidance that MCP Server throws away

**Evidence:**
```csharp
// ExcelDataModelTool.cs - OVERRIDES Core's suggestions
result.SuggestedNextActions =
[
    "Use 'list-measures' to see DAX measures",
    "Use 'list-relationships' to view table connections",
    "Use 'refresh' to update table data"
];

// Only uses SuggestedNextActions if Core didn't provide any
if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
{
    result.SuggestedNextActions = [
        "Use 'list-tables' to see Data Model structure",
        "Use 'create-measure' to add DAX calculations"
    ];
}
```

---

## Design Flaws

### 1. **Core Creates CLI-Specific Guidance**
- Core layer knows about CLI commands (`"Use 'dm-list-measures'"`)
- Violates layering - Core shouldn't know about CLI syntax

### 2. **Wasted Code in Core**
- 500+ lines of WorkflowGuidance classes in Core
- MCP Server **overrides** most of it
- Only CLI uses Core's guidance directly

### 3. **Maintenance Burden**
- Adding new feature = update Core guidance + MCP overrides
- Two places to maintain the same concept
- Easy to have inconsistencies

### 4. **Wrong Abstraction**
- `SuggestedNextActions: List<string>` is CLI-specific
- LLMs need **semantic hints**, not command strings
- MCP Server awkwardly shoehorns LLM guidance into CLI structure

### 5. **CLI vs. MCP Differences**

**CLI Needs:**
- Exact command syntax (`excelcli pq-refresh --query "Sales"`)
- Terminal workflow hints ("Run this next")
- Error recovery commands

**MCP Needs:**
- Action names (`action: "refresh"`)
- Semantic workflow context ("Query loaded to worksheet, not Data Model")
- Tool selection guidance ("Use excel_datamodel for DAX measures")

---

## Recommended Design

### Option 1: Push Workflow Guidance UP (Recommended)

**Move guidance to consumer layers where it belongs:**

```
Core (Business Logic Only)
├── No SuggestedNextActions
├── No WorkflowHint
├── Only operation results (Success, ErrorMessage, Data)
└── Focus: Pure Excel COM operations

CLI (Terminal User Guidance)
├── Creates SuggestedNextActions from result context
├── Knows CLI command syntax
├── Displays terminal-friendly hints
└── Example: "Run: excelcli pq-refresh --query Sales"

MCP Server (LLM Semantic Guidance)
├── Creates workflow hints from result context
├── Knows MCP tool names and actions
├── Provides semantic context for LLM reasoning
└── Example: "Query loaded to worksheet (not Data Model). Use excel_datamodel tool to add to Power Pivot."
```

**Changes Required:**
1. **Core:** Remove `SuggestedNextActions`, `WorkflowHint` from result classes
2. **Core:** Delete `PowerQueryWorkflowGuidance.cs`, `DataModelWorkflowGuidance.cs`
3. **CLI:** Create `CLI/WorkflowGuidance/` folder with CLI-specific guidance generators
4. **MCP Server:** Keep existing guidance logic (already correct)

**Benefits:**
- ✅ Core is pure business logic (no UI concerns)
- ✅ Each layer creates guidance appropriate for its audience
- ✅ No code duplication
- ✅ Clear separation of concerns
- ✅ Easier to test (Core results don't include presentation logic)

**Migration Effort:** Medium (remove from Core, add to CLI, MCP Server unchanged)

---

### Option 2: Keep Minimal Hints in Core

**Core provides semantic context, consumers format it:**

```csharp
// Core result class
public class PowerQueryResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    
    // NEW: Semantic context (not formatted strings)
    public WorkflowContext? Context { get; set; }
}

public class WorkflowContext
{
    public string OperationType { get; set; } // "import", "refresh", "delete"
    public bool IsConnectionOnly { get; set; }
    public string? LoadDestination { get; set; } // "worksheet", "datamodel", "both"
    public Dictionary<string, object> Metadata { get; set; } // Flexible context
}

// CLI interprets context
var suggestions = CliWorkflowGuidance.ForPowerQuery(result.Context);

// MCP Server interprets context
var hints = McpWorkflowGuidance.ForPowerQuery(result.Context);
```

**Benefits:**
- ✅ Core provides semantic information
- ✅ Consumers format appropriate for their audience
- ✅ Testable, structured data

**Drawbacks:**
- ⚠️ More complex result objects
- ⚠️ Still some workflow logic in Core

**Migration Effort:** High (redesign result classes)

---

### Option 3: Status Quo with Acknowledgment

**Keep current design, document that MCP Server overrides Core guidance:**

**Benefits:**
- ✅ No changes required

**Drawbacks:**
- ❌ Continued maintenance of unused Core guidance
- ❌ Confusing for developers
- ❌ Violates separation of concerns
- ❌ 500+ lines of dead code for MCP Server

---

## Recommendation: Option 1

**Rationale:**
1. **Clean Architecture** - Core is pure business logic
2. **No Wasted Code** - Each layer creates only what it needs
3. **Flexibility** - CLI and MCP can evolve independently
4. **Maintainability** - One place to update per consumer
5. **Testability** - Core results are pure data, easy to test

**Implementation Plan:**
1. Create `CLI/WorkflowGuidance/` folder
2. Move `PowerQueryWorkflowGuidance`, `DataModelWorkflowGuidance` to CLI layer
3. Update CLI commands to use new location
4. Remove `SuggestedNextActions`, `WorkflowHint` from Core result classes
5. MCP Server continues using existing guidance (already correct)
6. Update tests

**Estimated Effort:** 2-3 hours

---

## Current Guidance Usage Statistics

**Core Layer:**
- `SuggestedNextActions` assignments: ~50+ locations
- `WorkflowHint` assignments: ~40+ locations
- Workflow guidance classes: 3 files, ~500 lines total

**CLI Layer:**
- Uses `SuggestedNextActions`: 116 locations (all direct display)
- **100% dependency on Core guidance**

**MCP Server Layer:**
- Uses `SuggestedNextActions`: ~20 locations
- **~80% overrides Core guidance** with LLM-specific hints
- **20% uses Core guidance** only when Core doesn't provide any

---

## Questions for Decision

1. **Do we want Core to contain UI/presentation logic?**
   - Current answer: Yes (WorkflowGuidance classes exist)
   - Recommended: No (pure business logic)

2. **Should CLI and MCP Server share guidance logic?**
   - Current answer: Partially (Core provides base, MCP overrides)
   - Recommended: No (different audiences, different needs)

3. **Is it worth 2-3 hours to clean this up?**
   - Impact: -500 lines of Core code, cleaner architecture
   - Risk: Low (MCP Server already doesn't use Core guidance)

4. **Should we do this now or later?**
   - Now: Before adding more features that depend on flawed design
   - Later: Technical debt accumulates

---

## Conclusion

**Current design has Core creating CLI-specific guidance that MCP Server mostly ignores.**

**Recommended fix: Move workflow guidance to consumer layers (CLI, MCP Server) where it belongs.**

This creates clean separation of concerns and eliminates 500+ lines of mostly-unused code in Core.

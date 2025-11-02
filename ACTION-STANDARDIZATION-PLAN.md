# Action Name Standardization Plan

## Current State Analysis

### Inconsistent Patterns (BEFORE)

**"Get one item" actions:**
- PowerQuery: `View` (M code)
- Table: `Info` (table details)
- PivotTable: `GetInfo` (pivot details)
- NamedRange: `Get` (value)
- Connection: `View` (connection details)
- VBA: `View` (VBA code)
- DataModel: `ViewTable`, `ViewMeasure`
- Range: `GetRangeInfo`, `GetValues`, `GetFormulas`

**Problem:** 4 different names for the same concept!

## Standardized Verb System (AFTER)

### Core CRUD Verbs (Consistent Everywhere)

| Verb | Meaning | Example |
|------|---------|---------|
| `List` | Get all items | `excel_table(action: "List")` |
| `Get` | Get one item (data or metadata) | `excel_table(action: "Get", tableName: "Sales")` |
| `Create` | Make new item | `excel_table(action: "Create", ...)` |
| `Update` | Modify existing item | `excel_powerquery(action: "Update", ...)` |
| `Delete` | Remove item | `excel_table(action: "Delete", ...)` |
| `Rename` | Change item name | `excel_worksheet(action: "Rename", ...)` |
| `Copy` | Duplicate item | `excel_worksheet(action: "Copy", ...)` |

### Specialized Operation Verbs

| Verb | Meaning | Example |
|------|---------|---------|
| `Import` | Bring from external file | `excel_powerquery(action: "Import", sourcePath: "query.pq")` |
| `Export` | Save to external file | `excel_powerquery(action: "Export", targetPath: "query.pq")` |
| `Refresh` | Reload data from source | `excel_powerquery(action: "Refresh", ...)` |
| `Run` | Execute code | `excel_vba(action: "Run", ...)` |
| `Test` | Check connectivity | `excel_connection(action: "Test", ...)` |

### Property Management Verbs (Prefix-based)

| Pattern | Example |
|---------|---------|
| `Get[Property]` | `GetTabColor`, `GetVisibility`, `GetValidation` |
| `Set[Property]` | `SetTabColor`, `SetVisibility`, `SetStyle` |
| `Clear[Property]` | `ClearTabColor`, `ClearFilters`, `ClearFormats` |
| `Add[Property]` | `AddHyperlink`, `AddColumn`, `AddRowField` |
| `Remove[Property]` | `RemoveHyperlink`, `RemoveColumn`, `RemoveField` |

## Changes Required

### PowerQuery Tool
- Keep `View` (for M code source - different from Get which would return data)
- Keep `List`, `Create`, `Update`, `Delete`, `Import`, `Export`, `Refresh`
- `GetLoadConfig` → Keep (property pattern)
- `SetLoadToTable` → Keep (config pattern)
- `ListExcelSources` → Keep (already standardized)

**Verdict:** ✅ Already consistent

### Table Tool  
- `Info` → `Get` ✏️ **CHANGE**
- Keep `List`, `Create`, `Rename`, `Delete`
- Keep all property-based: `GetFilters`, `SetStyle`, `AddColumn`, etc.

### PivotTable Tool
- `GetInfo` → `Get` ✏️ **CHANGE**
- Keep `List`, `Create`, `Delete`, `Refresh`
- Keep field management: `AddRowField`, `RemoveField`, etc.

### DataModel Tool
- `ListTables` → Keep (clear what it lists)
- `ViewTable` → `GetTable` ✏️ **CHANGE**
- `ListMeasures` → Keep
- `ViewMeasure` → `Get` ✏️ **CHANGE** (when measureName specified)
- `GetModelInfo` → `GetInfo` ✏️ **CHANGE**
- Keep `Create`, `Update`, `Delete`, `Export`, `Refresh`

### Connection Tool
- Keep `View` (for connection definition - similar to PowerQuery.View)
- Keep `List`, `Import`, `Export`, `Test`, `Refresh`, `Delete`
- `GetProperties` → Keep (gets ALL properties object)
- `SetProperties` → Keep
- `UpdateProperties` → Keep (updates specific properties)

**Verdict:** ✅ Already consistent

### NamedRange Tool
- Keep `Get` (returns value)
- Keep `Set` (sets value)
- Keep `List`, `Create`, `Update`, `Delete`

**Verdict:** ✅ Already consistent

### VBA Tool
- Keep `View` (for VBA source code)
- Keep `List`, `Import`, `Export`, `Delete`, `Run`, `Update`

**Verdict:** ✅ Already consistent

### Worksheet Tool
- Keep `List`, `Create`, `Rename`, `Copy`, `Delete`
- Keep `GetTabColor`, `SetTabColor`, `ClearTabColor`
- Keep `GetVisibility`, `SetVisibility`, `Hide`, `VeryHide`, `Show`

**Verdict:** ✅ Already consistent

### Range Tool
- `GetRangeInfo` → `GetInfo` ✏️ **CHANGE** (for consistency)
- Keep `GetValues`, `SetValues`, `GetFormulas`, `SetFormulas`
- Keep `GetNumberFormats`, `SetNumberFormat`, `SetNumberFormats`
- Keep all other Get/Set/Clear/Add/Remove patterns

### File Tool
- Keep `CreateEmpty`, `Test`

**Verdict:** ✅ Already consistent

## Summary of Changes

### Minimal Changes Required (4 actions total)

1. **TableAction.Info** → **TableAction.Get**
2. **PivotTableAction.GetInfo** → **PivotTableAction.Get**
3. **DataModelAction.ViewTable** → **DataModelAction.GetTable**
4. **DataModelAction.ViewMeasure** → **DataModelAction.Get**
5. **DataModelAction.GetModelInfo** → **DataModelAction.GetInfo**
6. **RangeAction.GetRangeInfo** → **RangeAction.GetInfo**

## Rationale

**Why keep `View` for code/definitions?**
- `View` = Read-only inspection of SOURCE CODE/DEFINITION (M code, VBA, connection string)
- `Get` = Retrieve DATA or METADATA about an item
- Clear semantic distinction for LLMs

**Why use `Get` prefix for properties?**
- Standard pattern across all tools
- `GetTabColor`, `GetVisibility`, `GetFilters` - all consistent
- LLMs can predict: "To get X, use GetX action"

**Why `GetTable` instead of `Get` for DataModel?**
- DataModel has multiple item types (tables, measures, relationships)
- `GetTable` is explicit about WHAT you're getting
- Similar to `ListTables`, `ListMeasures` (explicit about item type)

## Implementation Order

1. Update enums in ToolActions.cs
2. Update ActionExtensions.cs mappings
3. Update MCP Server tool switch statements
4. Update Core Commands interfaces (if needed)
5. Update Core Commands implementations (if needed)
6. Update CLI commands (if any use these actions)
7. Update all documentation and prompts
8. Update tests

## Expected Benefits

✅ **Predictable API**: Learn once, apply everywhere  
✅ **Clear semantics**: `Get` = data/metadata, `View` = source code  
✅ **Easy discovery**: `Get[ItemType]` pattern makes it obvious  
✅ **Fewer questions**: LLMs don't have to ask "Is it Info or GetInfo or View?"  
✅ **Easier maintenance**: Consistent patterns = less cognitive overhead

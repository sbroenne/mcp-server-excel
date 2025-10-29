# Phase 2 Implementation Plan: Remove SuggestedNextActions from Core

## Current Status

✅ **Phase 0 Complete**: Shared validation layer with 22 action definitions across 3 domains, 32 passing tests

## Phase 2 Goal

**Remove ALL `SuggestedNextActions` from Core commands** - Core should return pure business data only.

## Scope Analysis

### PowerQueryCommands.cs
- **File size**: 2,890 lines
- **Instances**: 20+ occurrences of `SuggestedNextActions`
- **Methods affected**: UpdateAsync, ImportAsync, RefreshAsync, SetLoadToTableAsync, SetLoadToDataModelAsync, SetLoadToBothAsync, SetConnectionOnlyAsync, GetLoadConfigAsync

### Impact on Other Commands
- ParameterCommands: ~5 instances
- TableCommands: ~8 instances
- VbaCommands: ~10 instances
- FileCommands: ~3 instances
- ConnectionCommands: ~15 instances
- DataModelCommands: ~10 instances

**Total estimated**: ~70+ instances across 7 command files

## Implementation Strategy

### Option A: Complete Removal (Recommended)
**Approach**: Remove ALL SuggestedNextActions from Core in one comprehensive change

**Pros**:
- Clean break, no half-state
- Clear commit history
- Easier to reason about

**Cons**:
- Large PR (affects 7 files, ~70+ changes)
- Requires updating many tests
- Risk of breaking changes

**Estimated time**: 2-3 hours

### Option B: Incremental Removal
**Approach**: Remove one command at a time (PowerQuery → Parameter → Table → etc.)

**Pros**:
- Smaller PRs, easier review
- Can validate each step
- Lower risk per change

**Cons**:
- Multiple PRs to track
- Temporary inconsistency in codebase
- More total time

**Estimated time**: 1-2 days (7 separate PRs)

### Option C: Deprecation Path
**Approach**: Keep SuggestedNextActions but mark as deprecated, populate both old and new

**Pros**:
- Backward compatible
- Gradual migration

**Cons**:
- Contradicts our goal (Core should NOT know about presentation)
- Adds complexity
- Delays cleanup

**Estimated time**: Same as Option A, but delayed cleanup

## Recommended Approach: Option A (Complete Removal)

### Step 1: Remove from PowerQueryCommands
File: `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`

Methods to update:
- UpdateAsync (3 instances at lines 658, 669, 684)
- ImportAsync (5 instances)
- RefreshAsync (5 instances)  
- SetLoadToTableAsync (3 instances)
- SetLoadToDataModelAsync (2 instances)
- SetLoadToBothAsync (2 instances)
- SetConnectionOnlyAsync (1 instance)
- GetLoadConfigAsync (1 instance)

### Step 2: Remove from Other Commands
- ParameterCommands.cs
- TableCommands.cs
- VbaCommands.cs
- FileCommands.cs
- ConnectionCommands.cs
- DataModelCommands.cs

### Step 3: Update Result Models
**No changes needed** - Result classes already have the `SuggestedNextActions` property, we just stop populating it.

### Step 4: Update Tests
Find and update tests that assert on `SuggestedNextActions`:

```bash
# Find tests checking SuggestedNextActions
grep -r "SuggestedNextActions" tests/ --include="*.cs"
```

Expected test changes:
- Remove assertions on `SuggestedNextActions.Count`
- Remove assertions on specific suggestion strings
- Keep assertions on business data (Success, ErrorMessage, actual results)

### Step 5: Mark Property as Obsolete
In `ResultBase.cs`:

```csharp
/// <summary>
/// DEPRECATED: Will be removed in v2.0
/// Use ActionDefinitions for validation and suggestion generation in client layers
/// </summary>
[Obsolete("SuggestedNextActions will be removed in v2.0. Generate suggestions in MCP/CLI layers using ActionDefinitions.")]
public List<string> SuggestedNextActions { get; set; } = new();
```

## Testing Strategy

### Unit Tests
✅ Already passing - validation layer fully tested

### Integration Tests
Need to update:
- Remove assertions on `SuggestedNextActions`
- Ensure business logic assertions remain (Success, data returned, etc.)

### Manual Testing
After removal:
1. Build succeeds with 0 warnings
2. All integration tests pass
3. MCP Server still works (suggestions moved to Phase 3)
4. CLI still works (suggestions moved to Phase 4)

## Breaking Changes

**None** for v1.x - `SuggestedNextActions` property remains, just empty.

**Future (v2.0)**: Remove property entirely.

## Migration Path for Users

### For MCP Server (Phase 3)
```csharp
// Before (Core generates):
var result = await commands.ListAsync(batch);
// result.SuggestedNextActions populated

// After (MCP generates):
var result = await commands.ListAsync(batch);
var action = ActionDefinitions.PowerQuery.List;
var suggestions = GenerateMcpSuggestions(result, action);
// suggestions: structured JSON for LLM
```

### For CLI (Phase 4)
```csharp
// Before (Core generates):
var result = await commands.ListAsync(batch);
// Display result.SuggestedNextActions

// After (CLI generates):
var result = await commands.ListAsync(batch);
var action = ActionDefinitions.PowerQuery.List;
var suggestions = GenerateCliSuggestions(result, action);
// suggestions: human-friendly command examples
```

## Example: UpdateAsync Removal

**Before**:
```csharp
result.SuggestedNextActions = new List<string>
{
    "Query updated successfully, load configuration preserved",
    "Data automatically refreshed with new M code",
    "Use 'get-load-config' to verify configuration if needed"
};
```

**After**:
```csharp
// Just return result - no suggestions
return result;
```

**MCP will generate** (Phase 3):
```json
{
  "tool": "excel_powerquery",
  "action": "get-load-config",
  "requiredParams": { "excelPath": "file.xlsx", "queryName": "Sales" },
  "rationale": "Verify load configuration after update"
}
```

**CLI will generate** (Phase 4):
```
• Verify load configuration
  excelcli pq-get-load-config <file> "Sales"
```

## Risks & Mitigation

### Risk 1: Breaking existing MCP/CLI usage
**Mitigation**: Property stays, just empty. No breaking change in v1.x

### Risk 2: Lost workflow guidance
**Mitigation**: Phase 3/4 implement better suggestions using ActionDefinitions

### Risk 3: Test failures
**Mitigation**: Update tests incrementally, verify business logic still works

## Timeline

- **Phase 2**: 2-3 hours (Option A) or 1-2 days (Option B)
- **Phase 3**: 3-4 hours (MCP suggestion generation)
- **Phase 4**: 3-4 hours (CLI suggestion generation)
- **Phase 5**: 1 hour (documentation)

**Total remaining**: 1-2 days for complete implementation

## Next Steps

1. **Get approval** for Option A (complete removal) vs Option B (incremental)
2. **Execute Phase 2** - Remove all SuggestedNextActions from Core
3. **Verify builds and tests** - Ensure no regressions
4. **Proceed to Phase 3** - MCP Server generates own suggestions
5. **Proceed to Phase 4** - CLI generates own suggestions

## Success Criteria

✅ Core has ZERO references to CLI command names  
✅ Core has ZERO references to MCP action names  
✅ Core has ZERO `SuggestedNextActions` population  
✅ All tests pass  
✅ Build succeeds with 0 warnings  
✅ Property marked as `[Obsolete]` with clear migration message  

---

**Ready to proceed?** Awaiting decision on Option A vs Option B.

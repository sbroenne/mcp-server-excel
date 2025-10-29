# SuggestedNextActions Refactoring - Executive Summary

> **Status**: Design Complete, Ready for Implementation  
> **Issue**: [FEATURE] Review suggested next actions set-up  
> **Branch**: `copilot/design-next-actions-setup`

## Overview

This PR provides a comprehensive design and implementation plan for refactoring the `SuggestedNextActions` feature to address current brittleness, improve LLM effectiveness, and enhance human CLI user experience.

## Problem Statement

### Current Issues
1. **String-based implementation** - Hardcoded text scattered across 40+ files
2. **Duplication** - Same actions have different representations:
   - CLI: `"Use 'param-list' to see available parameters"`
   - MCP: `"Use 'list' to see available parameters"`
3. **No type safety** - Typos and wrong action names go undetected
4. **Maintenance burden** - Adding/changing actions requires updates in 3+ places
5. **Generic quality** - Many suggestions like "Check file exists" don't add value

### Impact
- High development friction when adding new commands
- Potential LLM confusion from inconsistent action names
- Poor user experience from generic suggestions
- Silent regression risk when refactoring actions

## Proposed Solution

### Architecture

```
NextAction (abstract base class)
├── ToMcp() → Structured JSON for LLM agents
├── ToCli() → Command examples for humans  
└── ToDescription() → Simple string (backward compatibility)

NextActionFactory (domain-specific builders)
├── PowerQuery.List(), View(), Import(), etc.
├── Parameter.List(), Get(), Set(), Create(), etc.
├── Table.List(), Info(), Rename(), etc.
└── ...

Concrete Implementations
├── ViewItemAction
├── ListItemsAction  
├── CreateItemAction
├── UpdateItemAction
└── ... (12 total action types)
```

### Key Benefits

#### For LLM Agents (MCP)
```json
{
  "tool": "excel_powerquery",
  "action": "view",
  "requiredParams": { "excelPath": "file.xlsx", "queryName": "Sales" },
  "rationale": "Inspect M code of 'Sales' query"
}
```
✅ Structured tool invocation ready  
✅ Parameter hints (required vs optional)  
✅ Workflow rationale for context  

#### For Human Users (CLI)
```
• View details of query 'Sales'
  excelcli pq-view <file> "Sales"
```
✅ Full command syntax with placeholders  
✅ Natural language descriptions  
✅ Copy-paste ready examples  

#### For Developers
```csharp
// Type-safe factory method
result.NextActions.Add(NextActionFactory.PowerQuery.View(queryName));

// Compiler catches typos and invalid actions
// Single source of truth for each action
// Refactoring-friendly (rename once, updates everywhere)
```
✅ Type safety (compiler validation)  
✅ Single source of truth  
✅ Refactoring-friendly  
✅ Testable  

## Implementation Strategy

### Phase 1: Core Infrastructure (1-2 days)
- Create base abstractions (`NextAction`, `NextActionType`, etc.)
- Implement concrete action classes (View, List, Create, etc.)
- Build `NextActionFactory` with domain builders
- Add `NextActions` property to `ResultBase`
- Mark `SuggestedNextActions` as `[Obsolete]`
- Write unit tests

### Phase 2: Core Commands (2-3 days)
- Migrate `PowerQueryCommands`
- Migrate `ParameterCommands`  
- Migrate `TableCommands`
- Migrate `DataModelCommands`
- Migrate `ScriptCommands` (VBA)
- Write integration tests

### Phase 3: MCP Server (1-2 days)
- Update all MCP tools to serialize `NextActions.ToMcp()`
- Test JSON serialization format
- Verify LLM agent compatibility

### Phase 4: CLI (1-2 days)  
- Update all CLI commands to display `NextActions.ToCli()`
- Test command formatting and examples
- Verify human readability

### Phase 5: Documentation (1 day)
- Update README with examples
- Document migration guide
- Mark deprecation timeline

**Total Estimate**: 6-10 days for complete migration

## Backward Compatibility

The design maintains full backward compatibility during migration:

```csharp
// Old code still works
if (result.SuggestedNextActions.Any())
{
    foreach (var suggestion in result.SuggestedNextActions)
    {
        Console.WriteLine(suggestion);
    }
}

// SuggestedNextActions auto-generates from NextActions
public List<string> SuggestedNextActions 
{
    get => NextActions.Select(a => a.ToDescription()).ToList();
}
```

Breaking change (removal of `SuggestedNextActions`) deferred to v2.0.

## Files Impact

### New Files (~15)
```
src/ExcelMcp.Core/Models/NextActions/
├── NextAction.cs
├── NextActionType.cs
├── NextActionMcp.cs
├── NextActionCli.cs
├── ViewItemAction.cs
├── ListItemsAction.cs
├── CreateItemAction.cs
├── UpdateItemAction.cs
├── DeleteItemAction.cs
├── RefreshItemAction.cs
├── ConfigureAction.cs
├── DiagnoseAction.cs
├── ImportAction.cs
├── ExportAction.cs
└── NextActionFactory.cs
```

### Modified Files (~25)
- `ResultTypes.cs` - Add `NextActions` property
- Core Commands (5 files) - Use `NextActionFactory`
- MCP Tools (5 files) - Serialize `ToMcp()`
- CLI Commands (5 files) - Display `ToCli()`
- Test files (10+) - New unit/integration tests

## Success Criteria

✅ **Type Safety**: No compilation with invalid action references  
✅ **DRY**: Each action defined once in factory  
✅ **Context-Aware**: Different suggestions based on operation state  
✅ **Dual Format**: Both MCP and CLI formats work correctly  
✅ **Backward Compatible**: Old code still works during migration  
✅ **Tested**: 80%+ code coverage for new classes  
✅ **Documented**: Examples and migration guide available  

## Testing Strategy

### Unit Tests
- Action serialization (ToMcp, ToCli, ToDescription)
- Factory method correctness
- Parameter handling
- Edge cases (null values, empty lists, etc.)

### Integration Tests  
- Core commands populate NextActions correctly
- MCP tools serialize valid JSON
- CLI tools display proper formatting
- Backward compatibility (SuggestedNextActions still works)

### Manual Verification
- Test with actual LLM agent (GitHub Copilot)
- Verify CLI output readability
- Check MCP server responses

## Documents Delivered

### 1. Design Document
**File**: `specs/suggested-next-actions-design.md`  
**Size**: 750+ lines

**Contents**:
- Problem analysis with examples from codebase
- Design principles (separate for MCP/CLI)
- Complete architecture diagrams
- Code examples for all components
- Migration path and timeline
- Example workflows (LLM and human)
- Open questions and recommendations

### 2. Implementation Guide  
**File**: `specs/suggested-next-actions-implementation-guide.md`  
**Size**: 500+ lines

**Contents**:
- Complete file organization
- Step-by-step implementation instructions
- Full code for all base classes
- Concrete action implementations
- Factory pattern examples
- Before/after migration examples
- Testing strategy with code samples
- Migration checklist
- Common patterns and anti-patterns

## Next Steps

### Recommended Approach
1. **Review** - Team reviews design and implementation guide
2. **Prototype** - Implement Phase 1 (core infrastructure)
3. **Proof of Concept** - Migrate one command (e.g., PowerQueryCommands.ListAsync)
4. **Validate** - Test MCP and CLI formatting with real usage
5. **Full Migration** - Complete remaining phases
6. **Documentation** - Update user-facing docs

### Alternative: Incremental Adoption
If full migration is too large:
1. Implement core infrastructure (Phase 1)
2. Use new system for **new commands only**
3. Gradually migrate existing commands as they're touched
4. Mark deprecation for v2.0 or v3.0

## Risk Assessment

### Low Risk ✅
- Backward compatible (no breaking changes in v1.x)
- Can be implemented incrementally
- Well-defined scope and boundaries
- Comprehensive test strategy

### Medium Risk ⚠️  
- Large surface area (40+ files affected)
- Requires coordination across 3 layers (Core/MCP/CLI)
- Learning curve for new pattern

### Mitigation Strategies
- Start with proof-of-concept (one command)
- Thorough testing at each phase
- Code review at phase boundaries
- Documentation and examples upfront

## Conclusion

This design provides a solid foundation for improving SuggestedNextActions while maintaining backward compatibility. The dual-format approach optimizes for both LLM agents and human users, and the factory pattern eliminates current brittleness.

**Recommendation**: Proceed with implementation, starting with Phase 1 proof-of-concept.

## Questions / Feedback

Please review the design documents and provide feedback on:
1. Architecture approach (factory pattern, base classes)
2. MCP format structure (is it optimal for LLMs?)
3. CLI format examples (are they clear for humans?)
4. Migration timeline (realistic? too aggressive?)
5. Testing strategy (comprehensive enough?)
6. Any missing use cases or edge cases

## References

- **Design Doc**: `specs/suggested-next-actions-design.md`
- **Implementation Guide**: `specs/suggested-next-actions-implementation-guide.md`
- **GitHub Issue**: [FEATURE] Review suggested next actions set-up
- **Branch**: `copilot/design-next-actions-setup`
- **MCP Spec**: https://modelcontextprotocol.io/

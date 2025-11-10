# PivotTable Strategy Pattern Refactor - Implementation Spec

## Current Problem
- Single `GetFieldForManipulation()` returns different types (CubeField vs PivotField) based on runtime detection
- `isOlap` flag scattered across 10+ methods with conditional logic
- Tests passing claimed in PR #162 but only 1/11 OLAP tests actually work
- Difficult to maintain and debug

## Proposed Solution: Strategy Pattern

### Architecture

```
IPivotTableFieldStrategy (interface)
├── RegularPivotTableFieldStrategy
│   ├── CanHandle() - checks pivot.PivotFields exists
│   ├── GetFieldForManipulation() - returns PivotField
│   └── AddRow/Column/Value/Filter/Remove/Set* methods
└── OlapPivotTableFieldStrategy
    ├── CanHandle() - checks pivot.CubeFields.Count > 0
    ├── GetFieldForManipulation() - returns CubeField + CreatePivotFields()
    └── AddRow/Column/Value/Filter/Remove/Set* methods (CubeField-specific)

PivotTableFieldStrategyFactory
└── GetStrategy(pivot) - selects OlapStrategy or RegularStrategy
```

### File Structure

```
src/ExcelMcp.Core/Commands/PivotTable/
├── IPivotTableFieldStrategy.cs (NEW) ✅ CREATED
├── PivotTableFieldStrategyFactory.cs (NEW) ✅ CREATED  
├── RegularPivotTableFieldStrategy.cs (NEW) - Extract from current code
├── OlapPivotTableFieldStrategy.cs (NEW) - OLAP-specific implementation
├── PivotTableCommands.cs (KEEP) - FindPivotTable, helpers
└── PivotTableCommands.Fields.cs (REFACTOR) - Delegate to strategies
```

### Implementation Steps

1. **RegularPivotTableFieldStrategy** - Extract existing logic
   - Copy current PivotTableCommands.Fields.cs methods
   - Remove all `isOlap` conditionals (keep regular path only)
   - Returns PivotField from GetFieldForManipulation
   - Uses pivot.PivotFields API exclusively

2. **OlapPivotTableFieldStrategy** - New OLAP implementation
   - Returns CubeField from GetFieldForManipulation
   - Calls CreatePivotFields() when PivotFields don't exist
   - Sets Orientation on CubeField (not PivotField!)
   - Uses pivot.CubeFields API exclusively
   - Skip data type detection (return "Cube")
   - Skip available values enumeration (not supported)

3. **Refactor PivotTableCommands.Fields.cs**
   - Each method becomes thin wrapper:
     ```csharp
     public async Task<PivotFieldResult> AddRowFieldAsync(...)
     {
         return await batch.Execute((ctx, ct) =>
         {
             var pivot = FindPivotTable(ctx.Book, pivotTableName);
             var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
             return strategy.AddRowField(pivot, fieldName, position, batch.WorkbookPath);
         });
     }
     ```
   - Remove GetFieldForManipulation from PivotTableCommands.cs (move to strategies)
   - Keep shared helpers: FindPivotTable, GetAreaName, GetComAggregationFunction, etc.

### Testing Strategy

**Existing Tests (Keep As-Is):**
- PivotTableCommandsTests.Fields.cs (28 tests) - Regular PivotTables
- PivotTableCommandsTests.OlapFields.cs (11 tests) - OLAP PivotTables

**Expected Results After Refactor:**
- Regular tests: 28/28 passing (unchanged behavior)
- OLAP tests: 11/11 passing (fix via OlapStrategy)

### Benefits

1. **Separation of Concerns** - OLAP logic isolated from Regular logic
2. **Type Safety** - Each strategy knows exactly what it returns
3. **Testability** - Can unit test each strategy independently
4. **Maintainability** - No more scattered `isOlap` conditionals
5. **Extensibility** - Easy to add new PivotTable types (PowerBI, SQL, etc.)
6. **SOLID Principles** - Single Responsibility, Open/Closed

### Risks & Mitigation

**Risk:** Large refactor might break existing tests
**Mitigation:** Implement RegularStrategy first, verify 28 tests pass, then OLAP

**Risk:** COM lifecycle issues with different object types
**Mitigation:** Each strategy owns its COM cleanup logic

**Risk:** Time to implement
**Mitigation:** ~2-3 hours for complete implementation + testing

### Success Criteria

- ✅ All 28 regular PivotTable tests pass
- ✅ All 11 OLAP PivotTable tests pass  
- ✅ Build with 0 warnings
- ✅ No `isOlap` conditionals in PivotTableCommands.Fields.cs
- ✅ Clear separation between Regular and OLAP logic

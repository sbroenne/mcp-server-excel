# Integration Test Coverage Analysis

**Generated:** 2025-01-XX  
**Purpose:** Identify missing integration tests for all commands

---

## Summary

| Command Class | Methods Implemented | Tests Exist | Coverage |
|--------------|---------------------|-------------|----------|
| RangeCommands | 43 methods | 26 tests | **60%** ⚠️ |
| TableCommands | 23 methods | 7 tests | **30%** ❌ |
| PivotTableCommands | 17 methods | 4 tests | **23%** ❌ |
| DataModelCommands | ~20 methods | ~15 tests | **75%** ✅ |
| PowerQueryCommands | ~12 methods | ~20 tests | **100%** ✅ |
| ConnectionCommands | ~10 methods | ~8 tests | **80%** ✅ |
| FileCommands | 3 methods | 3 tests | **100%** ✅ |
| ParameterCommands | ~8 methods | ~10 tests | **100%** ✅ |
| ScriptCommands | ~6 methods | ~8 tests | **100%** ✅ |
| SetupCommands | ~3 methods | ~3 tests | **100%** ✅ |
| SheetCommands | ~8 methods | ~6 tests | **75%** ✅ |

---

## Missing Test Coverage

### 🔴 RangeCommands (17 methods missing tests)

**Formatting Methods (NEW - Phase 1 implementation):**
- ❌ `FormatRangeAsync` - NO TESTS
- ❌ `AutoFitColumnsAsync` - NO TESTS
- ❌ `AutoFitRowsAsync` - NO TESTS
- ❌ `MergeCellsAsync` - NO TESTS
- ❌ `UnmergeCellsAsync` - NO TESTS
- ❌ `GetMergeInfoAsync` - NO TESTS
- ❌ `AddConditionalFormattingAsync` - NO TESTS
- ❌ `ClearConditionalFormattingAsync` - NO TESTS
- ❌ `SetCellLockAsync` - NO TESTS
- ❌ `GetCellLockAsync` - NO TESTS

**Validation Methods (NEW - Phase 1 implementation):**
- ❌ `ValidateRangeAsync` - NO TESTS
- ❌ `GetValidationAsync` - NO TESTS
- ❌ `RemoveValidationAsync` - NO TESTS

**Editing Methods (Partially tested):**
- ❌ `ClearFormatsAsync` - NO TESTS
- ❌ `CopyFormulasAsync` - NO TESTS
- ❌ `InsertCellsAsync` - NO TESTS
- ❌ `DeleteCellsAsync` - NO TESTS
- ❌ `InsertRowsAsync` - NO TESTS
- ❌ `DeleteRowsAsync` - NO TESTS
- ❌ `InsertColumnsAsync` - NO TESTS
- ❌ `DeleteColumnsAsync` - NO TESTS

**Hyperlinks (Partially tested):**
- ❌ `GetHyperlinkAsync` - NO TESTS (ListHyperlinksAsync tested)

**Existing Tests:**
- ✅ GetValuesAsync (5 tests)
- ✅ SetValuesAsync (3 tests)
- ✅ GetFormulasAsync (2 tests)
- ✅ SetFormulasAsync (2 tests)
- ✅ ClearAllAsync (1 test)
- ✅ ClearContentsAsync (2 tests)
- ✅ CopyAsync (1 test)
- ✅ CopyValuesAsync (1 test)
- ✅ FindAsync (1 test)
- ✅ ReplaceAsync (1 test)
- ✅ SortAsync (1 test)
- ✅ GetUsedRangeAsync (1 test)
- ✅ GetCurrentRegionAsync (1 test)
- ✅ GetRangeInfoAsync (1 test)
- ✅ AddHyperlinkAsync (1 test)
- ✅ RemoveHyperlinkAsync (1 test)
- ✅ ListHyperlinksAsync (1 test)
- ✅ GetNumberFormatsAsync (2 tests)
- ✅ SetNumberFormatAsync (4 tests)
- ✅ SetNumberFormatsAsync (2 tests)

---

### 🔴 TableCommands (16 methods missing tests)

**Lifecycle (Partially tested):**
- ✅ `ListAsync` - TESTED
- ✅ `CreateAsync` - TESTED
- ✅ `GetInfoAsync` - TESTED
- ❌ `RenameAsync` - NO TESTS
- ❌ `DeleteAsync` - NO TESTS
- ❌ `ResizeAsync` - NO TESTS

**Data Operations (NO tests):**
- ❌ `AppendRowsAsync` - NO TESTS
- ❌ `SetStyleAsync` - NO TESTS
- ❌ `ToggleTotalsAsync` - NO TESTS
- ❌ `SetColumnTotalAsync` - NO TESTS

**Data Model (NO tests):**
- ❌ `AddToDataModelAsync` - NO TESTS

**Filters (NO tests):**
- ❌ `ApplyFilterAsync` (criteria version) - NO TESTS
- ❌ `ApplyFilterAsync` (values version) - NO TESTS
- ❌ `ClearFiltersAsync` - NO TESTS
- ❌ `GetFiltersAsync` - NO TESTS

**Columns (NO tests):**
- ❌ `AddColumnAsync` - NO TESTS
- ❌ `RemoveColumnAsync` - NO TESTS
- ❌ `RenameColumnAsync` - NO TESTS

**Sorting (NO tests):**
- ❌ `SortAsync` (single column) - NO TESTS
- ❌ `SortAsync` (multiple columns) - NO TESTS

**Number Format (NEW - NO tests):**
- ❌ `GetColumnNumberFormatAsync` - NO TESTS
- ❌ `SetColumnNumberFormatAsync` - NO TESTS

**Structured References (Partially tested):**
- ✅ `GetStructuredReferenceAsync` - TESTED (4 tests)

---

### 🔴 PivotTableCommands (13 methods missing tests)

**Lifecycle (Partially tested):**
- ❌ `ListAsync` - NO TESTS
- ❌ `GetInfoAsync` - NO TESTS
- ✅ `CreateFromRangeAsync` - TESTED
- ✅ `CreateFromTableAsync` - TESTED
- ❌ `DeleteAsync` - NO TESTS
- ❌ `RefreshAsync` - NO TESTS

**Fields (Partially tested):**
- ✅ `ListFieldsAsync` - TESTED
- ✅ `AddRowFieldAsync` - TESTED
- ❌ `AddColumnFieldAsync` - NO TESTS
- ❌ `AddValueFieldAsync` - NO TESTS
- ❌ `AddFilterFieldAsync` - NO TESTS
- ❌ `RemoveFieldAsync` - NO TESTS
- ❌ `SetFieldFunctionAsync` - NO TESTS
- ❌ `SetFieldNameAsync` - NO TESTS
- ❌ `SetFieldFormatAsync` - NO TESTS

**Analysis (NO tests):**
- ❌ `GetDataAsync` - NO TESTS
- ❌ `SetFieldFilterAsync` - NO TESTS
- ❌ `SortFieldAsync` - NO TESTS

---

## Recommendations

### Priority 1: Range Formatting & Validation Tests (NEW Phase 1 features)

**Critical for spec compliance:**
1. Create `RangeCommandsTests.Formatting.cs` with tests for:
   - FormatRangeAsync (all format options)
   - MergeCellsAsync / UnmergeCellsAsync
   - GetMergeInfoAsync
   - AddConditionalFormattingAsync
   - ClearConditionalFormattingAsync
   - SetCellLockAsync / GetCellLockAsync

2. Create `RangeCommandsTests.AutoFit.cs` with tests for:
   - AutoFitColumnsAsync
   - AutoFitRowsAsync

3. Create `RangeCommandsTests.Validation.cs` with tests for:
   - ValidateRangeAsync (all validation types)
   - GetValidationAsync
   - RemoveValidationAsync

**Estimated:** 25-30 new tests (3-4 hours)

---

### Priority 2: Table Commands Coverage

**Create test files:**
1. `TableCommandsTests.Data.cs` - AppendRowsAsync, SetStyleAsync, ToggleTotalsAsync, SetColumnTotalAsync
2. `TableCommandsTests.Filters.cs` - ApplyFilterAsync (both), ClearFiltersAsync, GetFiltersAsync
3. `TableCommandsTests.Columns.cs` - AddColumnAsync, RemoveColumnAsync, RenameColumnAsync
4. `TableCommandsTests.Sort.cs` - SortAsync (single + multiple)
5. `TableCommandsTests.NumberFormat.cs` - GetColumnNumberFormatAsync, SetColumnNumberFormatAsync
6. Expand `TableCommandsTests.Lifecycle.cs` - RenameAsync, DeleteAsync, ResizeAsync

**Estimated:** 20-25 new tests (3-4 hours)

---

### Priority 3: PivotTable Commands Coverage

**Create test files:**
1. `PivotTableCommandsTests.Lifecycle.cs` - ListAsync, GetInfoAsync, DeleteAsync, RefreshAsync
2. `PivotTableCommandsTests.Fields.cs` - AddColumnFieldAsync, AddValueFieldAsync, AddFilterFieldAsync, RemoveFieldAsync, SetFieldFunctionAsync, SetFieldNameAsync, SetFieldFormatAsync
3. `PivotTableCommandsTests.Analysis.cs` - GetDataAsync, SetFieldFilterAsync, SortFieldAsync

**Estimated:** 15-20 new tests (2-3 hours)

---

### Priority 4: Range Editing Operations

**Expand `RangeCommandsTests.Editing.cs`:**
- ClearFormatsAsync
- CopyFormulasAsync
- InsertCellsAsync / DeleteCellsAsync
- InsertRowsAsync / DeleteRowsAsync
- InsertColumnsAsync / DeleteColumnsAsync

**Estimated:** 10-12 new tests (1-2 hours)

---

## Total Effort Estimate

- **Priority 1 (Range Formatting):** 25-30 tests, 3-4 hours
- **Priority 2 (Table Commands):** 20-25 tests, 3-4 hours
- **Priority 3 (PivotTable Commands):** 15-20 tests, 2-3 hours
- **Priority 4 (Range Editing):** 10-12 tests, 1-2 hours

**Total:** 70-87 new tests, 9-13 hours

---

## Current Status

**Well-Tested Commands:**
- ✅ PowerQueryCommands (100% coverage)
- ✅ FileCommands (100% coverage)
- ✅ ParameterCommands (100% coverage)
- ✅ ScriptCommands (100% coverage)
- ✅ SetupCommands (100% coverage)
- ✅ ConnectionCommands (80% coverage)
- ✅ DataModelCommands (75% coverage)

**Needs Attention:**
- ⚠️ RangeCommands (60% coverage, NEW features untested)
- ❌ TableCommands (30% coverage)
- ❌ PivotTableCommands (23% coverage)

---

## Next Steps

1. **Immediate:** Create Priority 1 tests for Range formatting/validation (Phase 1 spec compliance)
2. **Short-term:** Complete Priority 2 tests for Table commands
3. **Medium-term:** Complete Priority 3 tests for PivotTable commands
4. **Long-term:** Complete Priority 4 tests for Range editing operations

**Goal:** 95%+ coverage across all commands before next release.

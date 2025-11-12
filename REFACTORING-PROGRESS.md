# File Handle Refactoring Progress

> **Issue:** #173 - File locking with rapid sequential non-batch operations
> 
> **Solution:** FilePath-based API with FileHandleManager automatic handle caching
> 
> **Status:** In Progress (Phase 0-2 Complete, 7 interfaces remaining)

## Architecture Overview

**Old Pattern (Batch-Based):**
```csharp
await using var batch = await ExcelSession.BeginBatchAsync(filePath);
var result = await commands.MethodAsync(batch, args);
await batch.SaveAsync();
```

**New Pattern (FilePath-Based):**
```csharp
var result = await commands.MethodAsync(filePath, args);
await FileHandleManager.Instance.SaveAsync(filePath);
```

**Key Benefits:**
- Automatic handle caching by file path
- No file locking on sequential operations
- Simpler API (no batch context needed)
- Backward compatible during migration

---

## ✅ Phase 0: Foundation (COMPLETE)

**Commit:** d93d0c1

**Files Created:**
- `src/ExcelMcp.ComInterop/Session/ExcelWorkbookHandle.cs` - Wraps Excel COM objects with lifecycle management
- `src/ExcelMcp.ComInterop/Session/FileHandleManager.cs` - Thread-safe singleton with automatic caching
- `src/ExcelMcp.Core/Commands/IWorkbookCommands.cs` - Workbook lifecycle interface
- `src/ExcelMcp.Core/Commands/WorkbookCommands.cs` - Implementation using FileHandleManager

**Status:** ✅ Complete (5 methods)

---

## ✅ Phase 1: ISheetCommands (COMPLETE)

**Commits:**
- 537769e - Core + MCP Server
- 1054d4b - Core Tests
- d1daa0e - CLI

**Conversion:**
- ✅ Core: 13 filePath-based methods in `SheetCommands.FilePath.cs`
- ✅ Tests: 3 test files updated (Lifecycle, TabColor, Visibility)
- ✅ MCP Server: `ExcelWorksheetTool` converted to filePath API
- ✅ CLI: `SheetCommands.cs` converted to filePath API

**Methods:** List, Create, Rename, Copy, Delete, SetTabColor, GetTabColor, ClearTabColor, SetVisibility, GetVisibility, Show, Hide, VeryHide (13 total)

**Status:** ✅ Complete (13 methods)

---

## ✅ Phase 2: INamedRangeCommands (COMPLETE)

**Commits:**
- 38f6ee1 - Core + MCP Server
- 8e94501 - Core Tests + CLI

**Conversion:**
- ✅ Core: 7 filePath-based methods in `NamedRangeCommands.FilePath.cs`
- ✅ Tests: 3 test files updated (Lifecycle, Values, Validation)
- ✅ MCP Server: `ExcelNamedRangeTool` converted to filePath API
- ✅ CLI: `NamedRangeCommands.cs` converted to filePath API

**Methods:** List, Get, Set, Create, Update, Delete, CreateBulk (7 total)

**Status:** ✅ Complete (7 methods)

---

## ⏸️ Phase 3: IRangeCommands (DEFERRED - Architecture Decision Required)

**Reason for Deferral:**
IRangeCommands has complex COM operations that rely heavily on `IExcelBatch.Execute()` context pattern. Current FileHandleManager architecture doesn't provide equivalent batch context needed for:
- Dynamic range operations with COM object management
- Complex multi-step operations (copy/paste, find/replace, sort)
- Nested COM object access (Range → Cells → Value2)

**Methods (44 total):**
- Values: GetValues, SetValues, GetFormulas, SetFormulas, ClearContents, ClearAll, ClearFormats, ClearComments, ClearHyperlinks
- Copy/Paste: CopyRange, CutRange, PasteSpecial
- Insert/Delete: InsertRows, InsertColumns, DeleteRows, DeleteColumns
- Find/Replace: FindValue, ReplaceValue
- Sort: SortRange
- Discovery: GetUsedRange, GetCurrentRegion
- Hyperlinks: ListHyperlinks, CreateHyperlink, DeleteHyperlink, UpdateHyperlink
- Formatting: GetNumberFormat, SetNumberFormat, GetFont, SetFont, GetFill, SetFill, GetBorders, SetBorders
- Validation: GetValidation, SetValidation, ClearValidation
- Layout: AutoFitColumns, AutoFitRows, MergeRange, UnmergeRange
- Conditional: GetConditionalFormats, SetConditionalFormat, ClearConditionalFormats
- Protection: GetLocked, SetLocked

**Architecture Options:**
1. Add `ExecuteAsync<T>(Func<dynamic, dynamic, T>)` to ExcelWorkbookHandle (mimics batch context)
2. Refactor Range commands to use direct COM access (2000+ lines of duplication)
3. Keep Range commands batch-based until final cleanup phase

**Decision:** Deferred to later phase after simpler interfaces complete

**Status:** ⏸️ Deferred (44 methods pending architecture decision)

---

## ⏸️ Phase 4: ITableCommands (DEFERRED - Architecture Decision Required)

**Reason for Deferral:**
ITableCommands has similar architecture dependencies as IRangeCommands:
- Number formatting operations delegate to IRangeCommands (deferred)
- Filter/Sort operations require complex COM object manipulation
- ~23 methods total, ~40% depend on batch context pattern

**Methods (23 total):**
- Lifecycle: List, Get, Create, Rename, Delete, Resize
- Styling & Totals: SetStyle, ToggleTotals, SetColumnTotal
- Data: Append
- Data Model: AddToDataModel
- Filters: ApplyFilter (single), ApplyFilter (multiple), ClearFilters, GetFilters
- Columns: AddColumn, RemoveColumn, RenameColumn
- Structured References: GetStructuredReference
- Sort: Sort (single column), Sort (multiple columns)
- Number Format: GetColumnNumberFormat, SetColumnNumberFormat

**Decision:** Defer to later phase alongside IRangeCommands after architecture decision

**Status:** ⏸️ Deferred (23 methods pending architecture decision)

---

## ⏳ Remaining Phases

### Simplified Conversion Strategy

**Discovery:** Complex interfaces (PowerQuery, Connection, DataModel) have intricate helper methods and multi-step operations that require significant refactoring beyond simple FilePath conversion.

**New Approach:** Focus on simplest interfaces first to build momentum, then tackle complex ones with proper architecture.

### Phase 5: IVbaCommands (PRIORITY - Simplest)
- Methods: ~7 (List, Import, Export, Delete, Run, GetTrustStatus, SetTrustStatus)
- Complexity: Low-Medium (VBA trust configuration, but straightforward COM)
- Benefit: Small, self-contained, good learning case

### Phase 6: IQueryTableCommands
- Methods: ~8 (List, Create, Delete, Refresh, GetProperties, SetProperties)
- Complexity: Medium (QueryTable COM operations)

### Phase 7: IPowerQueryCommands (DEFERRED - Complex)
- Methods: ~18 (List, View, Import, Update, Export, Delete, Refresh, LoadTo, etc.)
- Complexity: High (Power Query M code management, load configurations, data model integration)
- **Issue:** Requires multiple helper methods (DetermineLoadConfiguration, ConfigureLoadDestinationAsync, RefreshQueryAsync, IsPowerQueryConnection, etc.) that have complex multi-step COM logic
- **Recommendation:** Defer until simpler interfaces complete

### Phase 8: IConnectionCommands (DEFERRED - Complex)
- Methods: ~15 (List, View, Import, Export, Update, Delete, Refresh, Test, Properties)
- Complexity: High (Connection string management, multiple connection types)
- **Recommendation:** Defer until simpler interfaces complete

### Phase 9: IDataModelCommands (DEFERRED - Complex)
- Methods: ~12 (ListTables, ListMeasures, ListRelationships, ExportMeasure, Refresh, Delete)
- Complexity: Very High (TOM API dependency, external assembly)
- **Recommendation:** Defer until simpler interfaces complete

### Phase 10: IPivotTableCommands (DEFERRED - Complex)
- Methods: ~15 (List, Create, Delete, Refresh, GetFields, AddField, RemoveField)
- Complexity: Very High (Complex COM PivotTable API, field management)
- **Recommendation:** Defer until simpler interfaces complete

---

## 📊 Summary Statistics

**Completed:**
- Phase 0: 5 methods (Workbook lifecycle)
- Phase 1: 13 methods (ISheetCommands)
- Phase 2: 7 methods (INamedRangeCommands)
- **Total: 25 methods converted** ✅

**Deferred:**
- Phase 3: 44 methods (IRangeCommands - architecture decision required) ⏸️
- Phase 4: 23 methods (ITableCommands - architecture decision required) ⏸️
- Phase 7-10: ~50 methods (Power Query, Connection, DataModel, PivotTable - complex helpers required) ⏸️

**Remaining (Simplified First):**
- Phase 5-6: ~15 methods (VBA, QueryTable - simpler interfaces) ⏳

**Grand Total:** ~167 methods to convert

**Current Progress:** 15% complete (25/167 methods, 117 deferred pending complexity/architecture resolution)

---

## Final Cleanup Phase

**After all command conversions complete:**
1. Remove batch infrastructure:
   - Delete `src/ExcelMcp.ComInterop/Session/ExcelSession.cs`
   - Delete `src/ExcelMcp.ComInterop/Session/ExcelBatch.cs`
   - Delete `src/ExcelMcp.ComInterop/Session/IExcelBatch.cs`
   - Delete `src/ExcelMcp.ComInterop/Session/ExcelContext.cs`
2. Remove batch-based method overloads from all interfaces
3. Remove `batchId` parameters from MCP tools
4. Remove batch CLI commands
5. Update documentation (README, MCP prompts, etc.)
6. Update `CORE-COMMANDS-AUDIT.md`

---

## Notes

- Existing batch API remains functional during migration
- Each interface converted end-to-end (Core → Tests → MCP → CLI)
- FileHandleManager automatically handles caching and cleanup
- No breaking changes until final cleanup phase

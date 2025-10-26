# REMAINING TEST CONVERSIONS - Quick Reference

## Completed (5 files, 48 methods, 56 errors fixed):
- CellCommandsTests (11 methods) ✅
- FileCommandsTests (10 methods) ✅  
- ParameterCommandsTests (9 methods) ✅
- SheetCommandsTests (10 methods) ✅
- ScriptCommandsTests (8 methods) ✅

## Build Status: 358 → 302 errors (56 fixed)

## Remaining Files (need batch conversion):
1. PowerQueryCommandsTests - ~25 methods
2. DataModelCommandsTests - ~12 methods
3. DataModelTomCommandsTests - ~5 methods
4. Connection + workflow tests - TBD

## Fast Conversion Pattern:
1. Add: using Sbroenne.ExcelMcp.Core.Session;
2. Fix constructor: CreateEmptyAsync().GetAwaiter().GetResult()
3. Convert methods: void → async Task, wrap in batch
4. Commit when file complete

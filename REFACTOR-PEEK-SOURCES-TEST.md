# Refactoring Plan: Move Peek/Sources/Test Actions to Appropriate Tools

## Problem Statement

From an LLM's perspective, the current architecture is confusing:

**Current (Confusing):**
- `excel_powerquery(action: "Peek", queryName: "X")` → Peeks at **Excel tables/named ranges**, NOT queries
- `excel_powerquery(action: "Sources")` → Lists **Excel tables/named ranges**, NOT queries  
- `excel_powerquery(action: "Test", queryName: "X")` → Tests **Excel tables/named ranges**, NOT queries

**LLM's Natural Assumption (WRONG!):**
- "I'll peek at the Power Query to see its data"
- "I'll list the source queries"
- "I'll test if the query works"

**What Actually Happens:**
- Error: "Source 'ConsumptionMilestones' not found" (when ConsumptionMilestones is a query!)
- LLM is confused because error says "Source" but LLM passed a query name
- No indication that "Peek" operates on tables/named ranges, not queries

## Root Cause

These actions were added to Power Query because Power Query M code references Excel sources via `Excel.CurrentWorkbook()`. But from a user's perspective, they belong with the data structures they operate on.

## Proposed Solution: Simplify and Reorganize (LLM-Optimized)

### Phase 1: Enhance Existing Actions (BREAKING CHANGE - OK per user)

#### 1. "Peek" → **REMOVE - Enhance "Info" instead**

**Rationale:** Info already checks if table/range exists. Just add preview data to Info response.

```csharp
// REMOVE: PowerQueryAction.Peek (redundant with enhanced Info)

// ENHANCE: TableAction.Info
// Add sampleData, headers to existing Info response
// TableInfoResult now includes preview (first 3-5 rows)

// ENHANCE: NamedRangeAction.Get  
// Already returns value - this IS the peek!
```

**Implementation:**
- **TableCommands.GetInfoAsync()**: Add headers + sample data to result
- **NamedRangeCommands.GetAsync()**: Already perfect - returns value
- **Remove** `PowerQueryCommands.PeekAsync()` entirely
- **Remove** `PowerQueryAction.Peek` enum value

#### 2. "Sources" → **Rename to "ListExcelSources"**

**Rationale:** Clarifies that it lists Excel tables/ranges, not Power Query sources.

```csharp
// CURRENT: Confusing name
PowerQueryAction.Sources  // "Sources of what? Queries?"

// NEW: Clear intent
PowerQueryAction.ListExcelSources  // "List Excel tables/ranges available to Power Query"
```

**Implementation:**
- Rename `PowerQueryCommands.SourcesAsync()` → `ListExcelSourcesAsync()`
- Rename `PowerQueryAction.Sources` → `PowerQueryAction.ListExcelSources`
- Update ActionExtensions mapping: `"sources"` → `"list-excel-sources"`

#### 3. "Test" → **REMOVE - Redundant with Info/Get**

**Rationale:** Testing if something exists is the same as getting info about it. If Info/Get succeeds, test passed!

```csharp
// REMOVE: PowerQueryAction.Test (redundant)
// USE INSTEAD:
//   - excel_table(action: "Info") → success means table exists
//   - excel_namedrange(action: "Get") → success means range exists
```

**Implementation:**
- **Remove** `PowerQueryCommands.TestAsync()` entirely
- **Remove** `PowerQueryAction.Test` enum value
- **Update docs**: "To test if table exists, use Info action"

### Phase 2: Update Documentation and Prompts

#### Update Tool Descriptions

**excel_table:**
```markdown
## Enhanced Action:
- **info**: Get table details INCLUDING preview data (headers, sample rows)

## Use Cases:
- Check if table exists AND see its structure: excel_table(action: "info", tableName: "Sales")
- Preview data before processing: Info now includes first 3-5 rows
- No need for separate "test" or "peek" actions - Info does it all!
```

**excel_namedrange:**
```markdown
## Existing Action (Already Perfect):
- **get**: Returns named range value - this IS the preview!

## Use Cases:
- Check if range exists AND get value: excel_namedrange(action: "get", namedRangeName: "StartDate")
- No need for separate "test" or "peek" - Get does it all!
```

**excel_powerquery:**
```markdown
## Renamed Action:
- **list-excel-sources** (formerly "sources"): List all Excel tables and named ranges available for Power Query to reference via Excel.CurrentWorkbook()

## Removed Actions (Use Instead):
- ❌ "peek" → Use excel_table(action: "info") or excel_namedrange(action: "get")
- ❌ "test" → Use excel_table(action: "info") or excel_namedrange(action: "get")

## Use Cases:
- Before writing M code, use list-excel-sources to see what tables/ranges are available
- Discover table names to use in: Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
```

### Phase 3: Update Error Messages

**Current (Confusing):**
```
peek failed for source 'ConsumptionMilestones' in 'file.xlsx': Source 'ConsumptionMilestones' not found
```

**New (Clear and Helpful):**
```
# If user tries obsolete actions:
Unknown action: Peek
The 'peek' action has been removed from excel_powerquery. Use instead:
  - excel_table(action: "info", tableName: "X") → Returns table details + preview data
  - excel_namedrange(action: "get", namedRangeName: "X") → Returns range value
  - excel_powerquery(action: "view", queryName: "X") → View Power Query M code

Unknown action: Test  
The 'test' action has been removed from excel_powerquery. Use instead:
  - excel_table(action: "info", tableName: "X") → Success = table exists
  - excel_namedrange(action: "get", namedRangeName: "X") → Success = range exists
```

## Implementation Checklist

### Code Changes

**PowerQuery (Simplify):**
- [ ] **PowerQueryCommands.Advanced.cs**: 
  - Remove `PeekAsync()` method entirely
  - Remove `TestAsync()` method entirely
  - Rename `SourcesAsync()` → `ListExcelSourcesAsync()`
- [ ] **ToolActions.cs**: 
  - Remove `PowerQueryAction.Peek`
  - Remove `PowerQueryAction.Test`
  - Rename `PowerQueryAction.Sources` → `PowerQueryAction.ListExcelSources`
- [ ] **ActionExtensions.cs**: Update mapping `"sources"` → `"list-excel-sources"`
- [ ] **ExcelPowerQueryTool.cs**: 
  - Remove switch cases for `Peek` and `Test`
  - Update switch case `Sources` → `ListExcelSources`
- [ ] **IPowerQueryCommands.cs**:
  - Remove `PeekAsync()` method signature
  - Remove `TestAsync()` method signature  
  - Rename `SourcesAsync()` → `ListExcelSourcesAsync()`

**Table (Enhance Info):**
- [ ] **TableCommands.cs**: Enhance `GetInfoAsync()` to include:
  - `Headers` list (first 10 column names)
  - `SampleData` (first 3-5 rows as List<List<object>>)
- [ ] **TableInfoResult.cs**: Add properties:
  - `public List<string> Headers { get; set; } = new();`
  - `public List<List<object?>> SampleData { get; set; } = new();`

**NamedRange (Already Perfect):**
- No changes needed! `Get` action already returns the value.

### Tests

**Remove:**
- [ ] **PowerQueryCommandsTests**: Remove all Peek tests
- [ ] **PowerQueryCommandsTests**: Remove all Test tests

**Update:**
- [ ] **PowerQueryCommandsTests**: Rename Sources tests → ListExcelSources tests
- [ ] **TableCommandsTests**: Update Info tests to verify Headers and SampleData in result

**Add:**
- [ ] **TableCommandsTests**: Add test for Info with empty table (0 rows)
- [ ] **TableCommandsTests**: Add test for Info with large table (verify only 3-5 sample rows returned)

### Documentation

**Update:**
- [ ] **excel_powerquery.md**: 
  - Remove Peek action
  - Remove Test action
  - Rename "sources" → "list-excel-sources"
  - Add migration notes
- [ ] **excel_table.md**: Document Info enhancement (includes preview data)
- [ ] **excel_namedrange.md**: Clarify Get action is the preview
- [ ] **tool_selection_guide.md**: Update patterns for checking table/range existence
- [ ] **COMMANDS.md**: Update CLI commands (if applicable)
- [ ] **README.md**: Update tool action counts

**Add:**
- [ ] **MIGRATION-V2.md**: Migration guide for breaking changes

### Migration Guide for Existing Users

Create `MIGRATION-V2.md`:

```markdown
# Breaking Change: Simplified API (v2.0.0)

## What Changed

**Removed from excel_powerquery (redundant actions):**
- ❌ `action: "Peek"` → Use enhanced Info/Get actions instead
- ❌ `action: "Test"` → Use Info/Get actions instead (success = exists)

**Renamed in excel_powerquery (clearer naming):**
- ⚠️ `action: "Sources"` → `action: "ListExcelSources"`

**Enhanced existing actions:**
- ✅ `excel_table(action: "Info")` now includes Headers + SampleData
- ✅ `excel_namedrange(action: "Get")` already returns value (no changes)

## Migration Guide

### Peek → Use Info or Get

**Before (Old - BROKEN):**
```javascript
// ❌ This will fail - Peek removed
excel_powerquery(action: "Peek", queryName: "TableName")
```

**After (New - CORRECT):**
```javascript
// ✅ Preview table data (now includes headers + sample rows)
excel_table(action: "Info", tableName: "TableName")
// Returns: { rowCount, columnCount, headers: [...], sampleData: [[...], [...]] }

// ✅ Preview named range value
excel_namedrange(action: "Get", namedRangeName: "StartDate")
// Returns: { value: "2024-01-01" }

// ✅ View Power Query M code (if you meant to peek at query)
excel_powerquery(action: "View", queryName: "ActualQueryName")
// Returns: { mCode: "let Source = ...", characterCount: 150 }
```

### Test → Use Info or Get (Success = Exists)

**Before (Old - BROKEN):**
```javascript
// ❌ This will fail - Test removed
excel_powerquery(action: "Test", queryName: "TableName")
```

**After (New - SIMPLER):**
```javascript
// ✅ Test if table exists (success = exists, error = not found)
try {
  await excel_table(action: "Info", tableName: "Sales");
  console.log("Table exists!");
} catch {
  console.log("Table not found");
}

// ✅ Test if named range exists
try {
  await excel_namedrange(action: "Get", namedRangeName: "StartDate");
  console.log("Range exists!");
} catch {
  console.log("Range not found");
}
```

### Sources → ListExcelSources

**Before (Old - DEPRECATED but still works in v2.0):**
```javascript
// ⚠️ Deprecated - will be removed in v3.0
excel_powerquery(action: "Sources")
```

**After (New - RECOMMENDED):**
```javascript
// ✅ Clear intent: List Excel tables/ranges available to Power Query
excel_powerquery(action: "ListExcelSources")
// Returns: { items: [{name: "Sales", type: "Table"}, {name: "StartDate", type: "NamedRange"}] }
```

## Why These Changes?

### Problem (v1.x)
- **Confusing**: "Peek at a Power Query" actually peeked at tables/ranges, not queries!
- **Redundant**: "Test" did the same thing as "Info" (check if exists)
- **Unclear naming**: "Sources" could mean many things

### Solution (v2.0)
- **Intuitive**: Actions match their data structures
  - `excel_table` → operates on tables
  - `excel_namedrange` → operates on ranges
  - `excel_powerquery` → operates on queries
- **Simpler**: Fewer actions to remember (Info does test + peek)
- **Clear naming**: "ListExcelSources" explicitly says what it lists

## LLM Benefits

As an LLM using this API, I now have:
✅ **Predictable patterns**: Info always checks existence + returns metadata
✅ **Clear errors**: "Table not found" (not "Source not found")  
✅ **Logical grouping**: Table operations in excel_table, range operations in excel_namedrange
✅ **Less confusion**: No more mixing query operations with table/range operations

## Upgrade Path

**v2.0.0 (Current):**
- Peek/Test removed (use Info/Get)
- Sources renamed to ListExcelSources (old name deprecated with warning)

**v3.0.0 (Future):**
- Sources deprecated name removed entirely (only ListExcelSources works)
```

## Benefits of This Refactoring

### For LLMs (like me!)

✅ **Fewer actions**: 2 actions removed (Peek, Test), 1 renamed (Sources)
✅ **Clearer intent**: Info = check existence + get metadata + preview
✅ **Better errors**: "Table 'X' not found" instead of "Source 'X' not found"
✅ **Logical API**: Each tool operates ONLY on its data structure
✅ **No confusion**: Can't accidentally "peek at a query" when I meant table

### For Users

✅ **Simpler workflow**: One action instead of two (no Test then Peek)
✅ **More info per call**: Info now includes preview data
✅ **Better discoverability**: Fewer actions = easier to learn
✅ **Clearer errors**: Error messages point to correct tool

### For Maintainers

✅ **Less code**: Remove PeekAsync/TestAsync entirely (leverage existing Info/Get)
✅ **Better separation**: PowerQuery only handles queries, not tables/ranges
✅ **Fewer tests**: Remove redundant test coverage
✅ **Clearer architecture**: Each tool has clear responsibilities

## Timeline

**Estimated effort: 3-4 hours total**

1. **Code changes** (1.5 hours)
   - Remove PeekAsync/TestAsync from PowerQueryCommands
   - Rename SourcesAsync → ListExcelSourcesAsync
   - Enhance TableCommands.GetInfoAsync() with Headers + SampleData
   - Update enums and switch statements
   
2. **Tests** (1 hour)
   - Remove PowerQuery Peek/Test tests
   - Update PowerQuery Sources tests
   - Enhance Table Info tests
   
3. **Documentation** (1 hour)
   - Update all prompt files
   - Update tool selection guide
   - Create migration guide
   
4. **PR review and merge** (30 minutes)

## Decision

**Proceed with refactoring?** ✅ YES - User confirmed breaking changes are acceptable

**Key improvements:**
- **Simpler**: Remove 2 redundant actions (Peek, Test)
- **Clearer**: Rename 1 action (Sources → ListExcelSources)  
- **Enhanced**: Info now includes preview data
- **Better DX**: LLMs can't accidentally "peek at a query" anymore

This will significantly improve LLM experience and make the API more intuitive.

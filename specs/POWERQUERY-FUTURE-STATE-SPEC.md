# Power Query Future State Specification

> **LLM-optimized API: Simplified, intuitive operations without artificial distinctions**

**Date**: 2025-01-29 (Updated: 2025-11-08)  
**Status**: PROPOSED FUTURE STATE  
**Based On**: LLM usage analysis + User feedback (2025-01-28: UpdateMCode footgun) + Excel COM API validation

## 🎯 Core Design Philosophy: No Artificial Distinctions for LLMs

**Problem**: The current API has artificial "atomic vs granular" distinctions that confuse LLMs:
- "Should I use `update-mcode` or `update-and-refresh`?"
- "Does update include refreshing data?"
- "Why are some operations atomic and others granular?"

**Solution**: Each operation does the **complete, intuitive thing**:
- ✅ `create` → Create query + load data (ONE complete operation)
- ✅ `update` → Update M code + refresh data (ONE complete operation)
- ✅ `load-to` → Change load destination + refresh (ONE complete operation)
- ✅ `refresh` → Refresh data only (when you only want to reload)
- ✅ `unload` → Remove data (make connection-only)

**Removed**: `update-and-refresh` (redundant), `update-mcode` (confusing name, incomplete operation)

---

## ⚠️ Critical Discovery: Excel COM API Limitations

**Investigation Results** (2025-01-29):

After validating the proposed API against Excel's COM API capabilities, two methods were found to be **unimplementable**:

### ErrorsAsync - ❌ NOT IMPLEMENTABLE

**Finding**: Excel WorkbookQuery object has **NO error properties or methods** in VBA.

**Evidence**: 
- [Official Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbookquery)
- Available properties: Application, Creator, Description, Formula (read-only), Name, Parent
- Available methods: Delete(), Refresh()
- **No error-related members exist**

**Current Implementation**: Returns placeholder message "No error information available through Excel COM interface"

**Impact**: Removed from proposed API (18 methods → 14 methods)

### EvalAsync - ⚠️ LIMITED FUNCTIONALITY

**Finding**: Can validate M code syntax but **cannot retrieve evaluation results**.

**Limitation**: WorkbookQuery.Refresh() validates syntax but doesn't expose computed values

**Solution**: Renamed to `ValidateSyntaxAsync()` with clear scope limitation

**Use Case**: Pre-flight syntax validation before creating permanent query (still useful for development workflows)

---

## 🚀 Quick Start for LLMs

> **Three essential patterns cover 95% of use cases**

### Pattern 1: Create Query with Data (Most Common)

```typescript
// Create query + load data in ONE operation
excel_powerquery({ 
  action: "create", 
  queryName: "Sales", 
  mCodePath: "sales.pq",
  loadDestination: "worksheet",  // Options: worksheet | data-model | both | connection-only
  targetSheet: "Data"
})
```

**Complete operation**: Imports M code, creates query, configures load destination, and loads data.

### Pattern 2: Update Existing Query

```typescript
// Update M code + refresh data automatically
excel_powerquery({ 
  action: "update", 
  queryName: "Sales", 
  mCodePath: "new-sales.pq" 
})
```

**Complete operation**: Updates M code AND refreshes data automatically. No separate refresh needed!

**Rationale**: When you update query logic, you want fresh data. The old "update without refresh" left stale data (footgun).

### Pattern 3: Change Where Data Loads

```typescript
// Change load destination + refresh
excel_powerquery({ 
  action: "load-to", 
  queryName: "Sales",
  loadDestination: "data-model",  // Switch from worksheet to data model
  targetSheet: null
})
```

**Complete operation**: Reconfigures where data loads AND refreshes to apply the change. 
  action: "set-load-destination",  // NEW: Apply load config without refreshing
  queryName: "ComplexQuery",
  loadTo: "DataModel"
})
excel_powerquery({ action: "refresh", queryName: "ComplexQuery" })  // Now load data
```

**Why Better**: Syntax validation catches errors early, connection-only creation separates definition from execution, explicit load control.

---

## Executive Summary

This specification proposes a **unified, LLM-optimized Power Query API** that eliminates current pain points discovered through diagnostic testing and real-world LLM usage patterns. The future state prioritizes:

1. **Atomic operations** - Single call accomplishes complete workflow
2. **Explicit intent** - Clear action names that match user goals
3. **Predictable behavior** - No hidden state, no surprise data loads
4. **Performance** - Eliminate redundant refresh operations
5. **Error resilience** - Handle Excel COM stress gracefully

---

## Current State Analysis

### Current API (18 methods)

From `IPowerQueryCommands.cs`:
```csharp
// CRUD Operations
ListAsync()
ViewAsync()
ImportAsync()           // Multi-parameter: loadDestination, worksheetName
UpdateAsync()           // Updates M code only
DeleteAsync()

// Refresh Operations
RefreshAsync()          // No loadDestination parameter
RefreshAsync(timeout)   // Overload with timeout

// Load Configuration (3 separate methods)
SetLoadToTableAsync()           // Atomic: config + refresh
SetLoadToDataModelAsync()       // Atomic: config + refresh
SetLoadToBothAsync()            // Atomic: config + refresh
SetConnectionOnlyAsync()        // Config only, no refresh

// Advanced Operations
LoadToAsync()           // Load connection-only to worksheet
ErrorsAsync()           // ❌ STUB - Excel COM API has no error properties
GetLoadConfigAsync()    // Read current config
ListExcelSourcesAsync() // Discovery
EvalAsync()             // ⚠️ LIMITED - Can validate syntax only, not retrieve results
```

### Current Pain Points (from LLM usage + diagnostics)

#### 1. **Inefficient Workflows** (LLM Confusion)

**Current "Common mistakes" from prompts**:
```markdown
❌ INEFFICIENT (data loaded twice):
1. import(loadDestination='connection-only')
2. refresh() 
3. set-load-to-table()  
```

**Why this happens**: LLMs think "connection-only" means "import without executing" but then discover they need to load data, leading to a 3-step workflow.

**Root Cause**: 
- `import(loadDestination='connection-only')` doesn't validate M code (no execution)
- `refresh()` creates temporary cache but no QueryTable
- `set-load-to-table()` creates QueryTable + triggers SECOND refresh

**Diagnostic Evidence** (Test 4 + Test 5):
- Connection-only queries have NO QueryTable (Queries.Count = 1, QueryTables.Count = 0)
- Creating QueryTable from connection-only triggers refresh automatically
- Result: **Double refresh** (refresh call + QueryTable creation)

#### 2. **Confusing Action Names** (Cognitive Load)

**LLM Perspective**: "I want to load data to a worksheet"

**Current Options**:
- `import(loadDestination='worksheet')` - Works but requires knowing parameter
- `set-load-to-table()` - Sounds like configuration, actually triggers refresh
- `LoadToAsync()` - Hidden in advanced section, not discoverable

**Problem**: 
- Action names don't match intent
- "set-load-to-X" sounds like configuration but is actually **atomic operation**
- LLMs default to `import()` then get confused by load behavior

#### 3. **Unexpected Refresh Behavior** (Hidden Magic)

**Current API**:
```csharp
// These methods say "Set" but actually REFRESH:
SetLoadToTableAsync()        // ATOMIC: config + refresh
SetLoadToDataModelAsync()    // ATOMIC: config + refresh
SetLoadToBothAsync()         // ATOMIC: config + refresh

// This method says "Set" and only configs (NO refresh):
SetConnectionOnlyAsync()     // Config only, no refresh
```

**LLM Confusion**:
1. "Set" methods have **inconsistent refresh behavior**
2. `SetConnectionOnlyAsync()` doesn't refresh (safe, no data load)
3. Other three `Set*` methods **DO refresh** (data load triggered)

**Diagnostic Evidence** (Test 1 + Test 2):
- QueryTable creation = automatic refresh
- Refresh doesn't create duplicate QueryTables (Excel maintains single QueryTable)
- M code updates work in isolation but can fail under Excel stress (Test 3)
- **LLM Insight**: I can implement retry logic myself - server should return clear errors, not hide them

#### 4. **Missing Granular Control** (Power User Needs)

**Current Gap**: Can't update M code WITHOUT triggering refresh on next operation

**Use Case**: User wants to update M code for 5 queries, then refresh all at once
```csharp
// DESIRED (not currently possible):
UpdateMCode("Query1", newCode);  // Update only, no refresh
UpdateMCode("Query2", newCode);  // Update only, no refresh
UpdateMCode("Query3", newCode);  // Update only, no refresh
RefreshAll();  // Single batch refresh

// CURRENT (forced individual refreshes):
UpdateAsync("Query1", newCode);  // May trigger refresh if QueryTable exists
UpdateAsync("Query2", newCode);  // Same issue
UpdateAsync("Query3", newCode);  // Same issue
```

**Diagnostic Evidence** (Test 3):
- M code updates work: `query.Formula = newMCode` succeeds
- Refresh required ONLY if QueryTable exists AND data needs updating
- UpdateAsync() could be split: update formula vs update + refresh

#### 5. **Defensive Code Waste** (Performance Impact)

**Current Pattern** (found in diagnostics):
```csharp
// ❌ UNNECESSARY: Excel doesn't create duplicates
while (queryTables.Count > 1) {
    queryTables.Item(queryTables.Count).Delete();
}
```

**Diagnostic Evidence** (Test 2):
- Excel maintains **single QueryTable** across unlimited refreshes
- NO duplicate QueryTables created on refresh
- Defensive cleanup code = wasted CPU cycles

---

## Diagnostic Discoveries Summary

### ✅ Confirmed Behaviors (Excel COM Ground Truth)

| Scenario | Finding | Impact on API Design |
|----------|---------|---------------------|
| **Load to Worksheet** | Creates 1 QueryTable, subsequent refreshes maintain same object | No cleanup needed |
| **Multiple Refreshes** | QueryTables.Count stays at 1 (no duplicates) | Remove defensive cleanup code |
| **M Code Update** | Works in isolation, can fail under Excel stress | Return clear errors with IsRetryable flag |
| **Connection-Only** | NO QueryTable created automatically | Can update M code freely |
| **Load Connection-Only** | Manual QueryTable creation triggers refresh | Explain double-load pattern |

### ❌ Invalidated Assumptions

1. ~~"Multiple refreshes create duplicate QueryTables"~~ → **FALSE**
2. ~~"M code updates incompatible with loaded queries"~~ → **FALSE** (works in isolation)
3. ~~"Connection-only queries need manual refresh after load"~~ → **PARTIALLY FALSE** (QueryTable creation auto-refreshes)

---

## Future State Proposal: Unified Power Query API

### Design Principles

1. **Explicit Intent Over Implementation** - Action names match user goals
2. **Atomic by Default** - Single operation = complete workflow
3. **Predictable Behavior** - No hidden refreshes, clear data load semantics
4. **Performance First** - Eliminate redundant operations
5. **Clear Error Reporting** - Return structured errors with retry metadata (LLMs handle retry logic)

### Proposed API (14 methods, -4 from current)

#### Core CRUD Operations (6 methods)

```csharp
/// <summary>
/// Lists all Power Query queries
/// </summary>
Task<PowerQueryListResult> ListAsync(IExcelBatch batch);

/// <summary>
/// Views M code for a query
/// </summary>
Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName);

/// <summary>
/// Creates new query from M code file
/// ATOMIC: Import M code + Load data to destination in ONE operation
/// DEFAULT: loadTo = QueryDestination.Worksheet (validate by executing)
/// </summary>
Task<PowerQueryCreateResult> CreateAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile,
    QueryDestination loadTo = QueryDestination.Worksheet,
    string? worksheetName = null);

/// <summary>
/// Updates ONLY the M code formula (no refresh)
/// Use RefreshAsync() separately if data update needed
/// </summary>
Task<OperationResult> UpdateMCodeAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile);

/// <summary>
/// Deletes query (removes both Query and QueryTable if exists)
/// </summary>
Task<OperationResult> DeleteAsync(
    IExcelBatch batch, 
    string queryName);
```

#### Data Load Operations (4 methods - Simplified & Explicit)

```csharp
/// <summary>
/// Loads query data to specified destination
/// ATOMIC: Set destination + Refresh data in ONE operation
/// Creates QueryTable if destination requires it (Worksheet, Both)
/// </summary>
Task<PowerQueryLoadResult> LoadToAsync(
    IExcelBatch batch,
    string queryName,
    QueryDestination destination,
    string? worksheetName = null);

/// <summary>
/// Refreshes query data from source
/// Applies to queries already loaded (has QueryTable or in Data Model)
/// Connection-only queries: Use LoadToAsync() first
/// </summary>
Task<PowerQueryRefreshResult> RefreshAsync(
    IExcelBatch batch, 
    string queryName,
    TimeSpan? timeout = null);

/// <summary>
/// Changes query to connection-only (removes QueryTable, keeps M code)
/// No refresh triggered (safe operation)
/// </summary>
Task<OperationResult> UnloadAsync(
    IExcelBatch batch, 
    string queryName);

/// <summary>
/// Gets current load configuration
/// Returns: Destination, WorksheetName (if applicable), HasData
/// NOTE: Action name "get-load-config" follows query pattern for consistency with "set-load-destination"
/// </summary>
Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(
    IExcelBatch batch, 
    string queryName);
```

**Action Naming Consistency Note**:
- `view` → Read entire query (M code + all metadata)
- `get-load-config` → Read specific property (pairs with `set-load-destination`)
- Both patterns valid: whole-object read vs property-specific read

#### Advanced Operations (4 methods)

```csharp
/// <summary>
/// Lists available data sources (Excel.CurrentWorkbook() sources)
/// </summary>
Task<WorksheetListResult> ListExcelSourcesAsync(IExcelBatch batch);

/// <summary>
/// Validates M code syntax by creating temporary query + refreshing
/// LIMITATION: Can only validate syntax, cannot retrieve evaluation results
/// (Excel COM API doesn't expose query evaluation results, only success/failure)
/// Use Case: Syntax validation before creating permanent query
/// </summary>
Task<PowerQueryEvalResult> ValidateSyntaxAsync(
    IExcelBatch batch, 
    string mExpression);

/// <summary>
/// Updates M code AND refreshes data in ONE operation
/// Convenience method combining UpdateMCodeAsync + RefreshAsync
/// </summary>
Task<PowerQueryRefreshResult> UpdateAndRefreshAsync(
    IExcelBatch batch,
    string queryName,
    string mCodeFile,
    TimeSpan? timeout = null);

/// <summary>
/// Refreshes all queries in workbook
/// </summary>
Task<PowerQueryRefreshAllResult> RefreshAllAsync(
    IExcelBatch batch,
    TimeSpan? timeout = null);
```

> **⚠️ COM API Limitations Discovered**
> 
> The following methods are **NOT INCLUDED** due to Excel COM API constraints:
> 
> **ErrorsAsync** - ❌ REMOVED  
> - **Why**: WorkbookQuery object has NO error properties/methods in Excel VBA
> - **Evidence**: [Official docs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbookquery) show only: Application, Creator, Description, Formula (read-only), Name, Parent, Delete(), Refresh()
> - **Current Implementation**: Returns hardcoded "No error information available through Excel COM interface"
> - **Alternative**: Errors visible in Excel UI only; LLMs should handle refresh failures through try-catch
> 
> **EvalAsync** - ⚠️ RENAMED to `ValidateSyntaxAsync` (limited scope)  
> - **Why**: Can create temp query and call Refresh() to validate syntax, but cannot retrieve actual evaluation results
> - **Limitation**: WorkbookQuery has no properties to access computed values
> - **Use Case**: Syntax validation before creating permanent query (useful but limited)

### New Enum: QueryDestination

```csharp
/// <summary>
/// Specifies where Power Query data should be loaded
/// </summary>
public enum QueryDestination
{
    /// <summary>
    /// Load to worksheet as QueryTable (visible to users)
    /// Creates QueryTable, data appears in Excel
    /// NOT in Data Model (can't use DAX)
    /// </summary>
    Worksheet,

    /// <summary>
    /// Load to Power Pivot Data Model only
    /// Data available for DAX measures, NOT visible in worksheet
    /// </summary>
    DataModel,

    /// <summary>
    /// Load to BOTH worksheet AND Data Model
    /// Data visible to users AND available for DAX
    /// </summary>
    Both,

    /// <summary>
    /// Connection-only: M code imported but NOT executed
    /// No data loaded, no validation
    /// Use LoadToAsync() when ready to load data
    /// </summary>
    ConnectionOnly
}
```

### API Comparison: Current vs Future

| **Operation** | **Current API** | **Future API** | **Change** |
|---------------|-----------------|----------------|------------|
| Create query + load to worksheet | `ImportAsync(name, file, loadDestination='worksheet', sheet)` | `CreateAsync(name, file, QueryDestination.Worksheet, sheet)` | Renamed: Import → Create |
| Create query without loading | `ImportAsync(name, file, loadDestination='connection-only')` | `CreateAsync(name, file, QueryDestination.ConnectionOnly)` | Same behavior, clearer enum |
| Update M code only | `UpdateAsync(name, file)` ⚠️ May refresh | `UpdateMCodeAsync(name, file)` ✅ Never refreshes | Explicit: no side effects |
| Update M code + refresh | N/A (manual 2-step) | `UpdateAndRefreshAsync(name, file)` | New convenience method |
| Load to worksheet | `SetLoadToTableAsync(name)` ❌ Confusing name | `LoadToAsync(name, QueryDestination.Worksheet)` | Clear intent |
| Load to data model | `SetLoadToDataModelAsync(name)` ❌ Confusing name | `LoadToAsync(name, QueryDestination.DataModel)` | Clear intent |
| Load to both | `SetLoadToBothAsync(name, sheet)` ❌ Confusing name | `LoadToAsync(name, QueryDestination.Both, sheet)` | Clear intent |
| Unload data | `SetConnectionOnlyAsync(name)` ❌ Confusing name | `UnloadAsync(name)` | Clear intent: removes QueryTable |
| Refresh data | `RefreshAsync(name)` | `RefreshAsync(name)` | Unchanged |
| Refresh + change destination | `RefreshAsync(name, loadDest, sheet)` ⚠️ Hidden | `LoadToAsync(name, dest, sheet)` | Explicit: load = atomic operation |

### Breaking Changes Summary

| **Change Type** | **Count** | **Impact** |
|-----------------|-----------|-----------|
| Method Renamed | 5 | `ImportAsync` → `CreateAsync`, `SetLoad*` → `LoadToAsync`, `SetConnectionOnly` → `UnloadAsync` |
| Method Split | 1 | `UpdateAsync` → `UpdateMCodeAsync` (no refresh) + `UpdateAndRefreshAsync` (with refresh) |
| Parameter Changed | 1 | `loadDestination: string` → `loadTo: QueryDestination` enum |
| Method Removed | 3 | `SetLoadToTableAsync`, `SetLoadToDataModelAsync`, `SetLoadToBothAsync` (replaced by `LoadToAsync`) |
| Method Added | 2 | `UpdateAndRefreshAsync`, `RefreshAllAsync` |

**Migration Difficulty**: MEDIUM (rename + parameter change, but 1:1 mapping exists)

---

## Error Handling Philosophy: LLMs Handle Retries

### Why Server Should NOT Retry

**Traditional Approach** (Server-side retry):
```csharp
// ❌ DON'T DO THIS: Server implements retry logic
public async Task<OperationResult> UpdateWithRetry(...)
{
    for (int attempt = 0; attempt < 3; attempt++)
    {
        try
        {
            await UpdateAsync(...);
            return new OperationResult { Success = true };
        }
        catch (COMException ex) when (ex.HResult == 0x800706BE)
        {
            if (attempt == 2) throw;
            await Task.Delay(1000 * (attempt + 1));
        }
    }
}
```

**Problems**:
1. ⚠️ **Hidden behavior** - User (LLM) doesn't know retry is happening
2. ⚠️ **Fixed strategy** - Can't adjust retry logic per use case
3. ⚠️ **Blocking** - Ties up server resources during retry delays
4. ⚠️ **No context** - Server doesn't know if user wants retries or fast failure

### LLM-First Approach (Recommended)

**Server Responsibility**: Return **clear, structured error information**
```csharp
// ✅ CORRECT: Server returns error metadata
return new OperationResult
{
    Success = false,
    ErrorMessage = "Excel COM timeout (0x800706BE)",
    ErrorCategory = "ExcelStress",
    IsRetryable = true,
    RetryGuidance = "Exponential backoff recommended",
    SuggestedNextActions = new[] { "Retry with delay", "Close/reopen workbook" }
};
```

**LLM Responsibility**: Orchestrate retries based on context
```typescript
// ✅ LLM implements retry logic with full context
async function updateQuerySmart(queryName: string, codeFile: string) {
    // LLM knows: This is critical update, retry 5 times
    for (let attempt = 0; attempt < 5; attempt++) {
        const result = await excel_powerquery({
            action: 'update-mcode',
            queryName,
            sourcePath: codeFile
        });
        
        if (result.success) {
            return result;
        }
        
        // LLM decides: Is this retryable?
        if (!result.isRetryable) {
            throw new Error(`Non-retryable: ${result.errorMessage}`);
        }
        
        // LLM implements strategy: Exponential backoff
        const delay = Math.min(1000 * Math.pow(2, attempt), 10000);
        console.log(`Retry ${attempt + 1}/5 after ${delay}ms...`);
        await sleep(delay);
    }
    
    // LLM decides: Try alternative approach
    console.log('Max retries exceeded, trying alternative...');
    await closeAndReopenWorkbook();
    return await excel_powerquery({ action: 'update-mcode', ... });
}
```

### Benefits of LLM-Orchestrated Retries

| **Aspect** | **Server Retry** | **LLM Retry** |
|------------|------------------|---------------|
| **Flexibility** | ❌ Fixed strategy | ✅ Context-aware (critical vs non-critical) |
| **User Visibility** | ❌ Hidden from user | ✅ LLM can explain what's happening |
| **Resource Usage** | ❌ Blocks server thread | ✅ LLM controls timing |
| **Failure Handling** | ❌ Generic "max retries exceeded" | ✅ LLM tries alternatives (close/reopen) |
| **Batch Operations** | ❌ Each operation retries independently | ✅ LLM batches retries intelligently |

### Real-World Example: Batch Query Updates

**Scenario**: Update M code for 10 queries

**Server Retry Approach** (inefficient):
```csharp
// Each update retries independently (up to 30 operations)
for (int i = 0; i < 10; i++)
{
    await UpdateWithRetry(queries[i], code[i]);  // 1-3 attempts each
}
```

**LLM Retry Approach** (intelligent):
```typescript
// LLM batches operations and retries strategically
const results = await Promise.all(
    queries.map(q => updateQuerySmart(q.name, q.code))
);

// Check for failures
const failures = results.filter(r => !r.success);

if (failures.length > 0 && failures.every(f => f.isRetryable)) {
    // LLM decides: Excel is stressed, give it a break
    console.log('Excel under stress, waiting 5 seconds...');
    await sleep(5000);
    
    // Retry only failures
    const retryResults = await Promise.all(
        failures.map(f => updateQuerySmart(f.queryName, f.codeFile))
    );
}
```

### Design Principle

> **Server: Return structured errors with IsRetryable metadata**  
> **LLM: Implement retry logic based on operation context**

This separation of concerns gives LLMs the **flexibility to make smart decisions** while keeping the server **simple and predictable**.

---

## LLM Usage Optimization

### Before: Inefficient Workflow (Current)

**LLM Intent**: "Create a Power Query from a file and load it to a worksheet"

**Current Steps** (LLM discovers through trial & error):
```typescript
// ❌ INEFFICIENT: LLM tries connection-only first (no validation = scary)
1. excel_powerquery(action: 'import', queryName: 'Sales', 
                    sourcePath: 'sales.pq', 
                    loadDestination: 'connection-only')
   → M code imported but NOT executed (no validation)
   
// LLM realizes data not loaded, tries refresh
2. excel_powerquery(action: 'refresh', queryName: 'Sales')
   → Data refreshed to temporary cache (no QueryTable yet)
   
// LLM realizes data not visible, tries set-load-to-table
3. excel_powerquery(action: 'set-load-to-table', queryName: 'Sales')
   → Creates QueryTable + SECOND REFRESH (double data load!)
```

**Total Operations**: 3 calls, 2 refreshes (inefficient)

### After: Efficient Workflow (Future)

**LLM Intent**: "Create a Power Query from a file and load it to a worksheet"

**Future Steps** (clear & atomic):
```typescript
// ✅ EFFICIENT: Single atomic operation
1. excel_powerquery(action: 'create', queryName: 'Sales',
                    sourcePath: 'sales.pq',
                    loadTo: 'Worksheet')
   → M code imported + QueryTable created + Data loaded in ONE operation
```

**Total Operations**: 1 call, 1 refresh (optimal)

### Common LLM Workflows (Optimized)

#### Workflow 1: Import & Validate Data

**Current** (3 steps):
```typescript
1. import(loadDestination='connection-only')  // No validation
2. refresh()                                   // Validate but no QueryTable
3. set-load-to-table()                         // Create QueryTable + refresh again
```

**Future** (1 step):
```typescript
1. create(loadTo='Worksheet')  // Import + validate + load in one call
```

#### Workflow 2: Load to Data Model for DAX

**Current** (correct but confusing action name):
```typescript
1. import(loadDestination='data-model')  // OR
1. set-load-to-data-model()              // Confusing: "set" sounds passive
```

**Future** (clear intent):
```typescript
1. create(loadTo='DataModel')  // Clear: creating + loading
1. load-to(destination='DataModel')  // Clear: actively loading
```

#### Workflow 3: Update M Code Without Refresh

**Current** (unclear if refresh happens):
```typescript
1. update(queryName, newCode)  // ⚠️ May refresh if QueryTable exists
```

**Future** (explicit control):
```typescript
1. update-mcode(queryName, newCode)  // ✅ Never refreshes
1. refresh(queryName)                 // ✅ Explicit refresh when ready
```

#### Workflow 4: Batch Update Multiple Queries

**Current** (forced individual refreshes):
```typescript
1. update('Query1', code1)  // May refresh
2. update('Query2', code2)  // May refresh
3. update('Query3', code3)  // May refresh
```

**Future** (explicit batch control):
```typescript
1. update-mcode('Query1', code1)  // No refresh
2. update-mcode('Query2', code2)  // No refresh
3. update-mcode('Query3', code3)  // No refresh
4. refresh-all()                   // Single batch refresh
```

---

## MCP Tool Design (Future)

### Proposed Actions (15 total, -3 from current)

```typescript
excel_powerquery(action: string, ...)

Actions:
  // Core CRUD (6)
  - 'list'              // List all queries
  - 'view'              // View M code
  - 'create'            // Import M code + load data (atomic)
  - 'update-mcode'      // Update M code only (no refresh)
  - 'export'            // Export M code to file
  - 'delete'            // Delete query
  
  // Data Load (4)
  - 'load-to'           // Load data to destination (atomic: config + refresh)
  - 'refresh'           // Refresh existing query data
  - 'unload'            // Remove QueryTable (connection-only mode)
  - 'get-load-config'   // Get current load state
  
  // Advanced (5)
  - 'errors'            // Show execution errors
  - 'sources'           // List available data sources
  - 'eval'              // Evaluate M code interactively
  - 'update-and-refresh' // Update M code + refresh (atomic)
  - 'refresh-all'       // Refresh all queries
```

### Parameter Design

```typescript
// Unified loadTo parameter (replaces loadDestination string)
loadTo: 'Worksheet' | 'DataModel' | 'Both' | 'ConnectionOnly'

// Example calls:
excel_powerquery({
  action: 'create',
  queryName: 'Sales',
  sourcePath: 'sales.pq',
  loadTo: 'Worksheet',       // ✅ Clear: load to worksheet
  worksheetName: 'SalesData' // Optional custom sheet name
})

excel_powerquery({
  action: 'load-to',
  queryName: 'Sales',
  destination: 'DataModel'   // ✅ Clear: loading to data model
})

excel_powerquery({
  action: 'update-mcode',
  queryName: 'Sales',
  sourcePath: 'sales-v2.pq'  // ✅ Clear: M code only, no refresh
})

excel_powerquery({
  action: 'update-and-refresh',
  queryName: 'Sales',
  sourcePath: 'sales-v2.pq'  // ✅ Clear: M code + refresh atomic
})
```

### LLM Prompt Guidance (Future)

```markdown
## excel_powerquery Tool

**Actions**: list, view, create, update-mcode, export, delete, load-to, refresh, unload, 
            get-load-config, errors, sources, eval, update-and-refresh, refresh-all

**When to use excel_powerquery**:
- External data sources (databases, APIs, files)
- Power Query M code transformations
- Use excel_table for data already in Excel
- Use excel_datamodel for DAX measures after loading to data model

**Action guide**:
- create: Import M code + load data in ONE operation (use loadTo parameter)
- update-mcode: Update M code ONLY (no refresh, safe for batch updates)
- update-and-refresh: Update M code + refresh in ONE operation (convenience)
- load-to: Change where data loads + refresh (atomic operation)
- refresh: Refresh existing loaded query (must already have QueryTable or be in Data Model)
- unload: Remove QueryTable, keep M code (connection-only mode)

**loadTo parameter values**:
- 'Worksheet': Load to worksheet (visible to users, NOT in Data Model)
- 'DataModel': Load to Power Pivot (ready for DAX, NOT visible in worksheet)
- 'Both': Load to worksheet AND Data Model (best of both worlds)
- 'ConnectionOnly': Import M code but don't execute/validate

**Common workflows**:
✅ Create + validate: create(loadTo='Worksheet')
✅ Load to Data Model: create(loadTo='DataModel')
✅ Update without refresh: update-mcode() then refresh() when ready
✅ Batch update: update-mcode × N, then refresh-all()

**Avoid these patterns**:
❌ create(loadTo='ConnectionOnly') + refresh() + load-to('Worksheet')
   → Double refresh! Use create(loadTo='Worksheet') instead
   
❌ update-mcode() × 5 + refresh() × 5
   → Use update-mcode() × 5 + refresh-all() instead
```

---

## Implementation Plan

### Phase 1: Core Refactoring (No Breaking Changes)

**Goal**: Prepare codebase for future API without breaking existing code

**Changes**:
1. ✅ Remove defensive QueryTable cleanup code (Test 2 evidence)
2. ✅ Return structured error info for Excel stress (IsRetryable, ErrorCategory)
3. ✅ Document connection-only behavior (Test 4 + Test 5 evidence)
4. ✅ Add `QueryDestination` enum (internal use only, not exposed yet)
5. ✅ Create internal helper methods:
   - `CreateQueryWithDestinationAsync()` (future `CreateAsync`)
   - `LoadQueryToDestinationAsync()` (future `LoadToAsync`)
   - `UpdateMCodeOnlyAsync()` (future `UpdateMCodeAsync`)

**Timeline**: 1-2 weeks
**Risk**: LOW (no breaking changes, only internal improvements)

### Phase 2: New Methods (Additive Only)

**Goal**: Add future API methods alongside existing ones

**Changes**:
1. ✅ Add `CreateAsync()` - new name for `ImportAsync()`
2. ✅ Add `UpdateMCodeAsync()` - update without refresh
3. ✅ Add `UpdateAndRefreshAsync()` - convenience method
4. ✅ Add `LoadToAsync()` - unified load destination method
5. ✅ Add `UnloadAsync()` - clearer name for `SetConnectionOnlyAsync()`
6. ✅ Add `RefreshAllAsync()` - batch refresh
7. ✅ Mark old methods `[Obsolete]` with migration guidance

**Timeline**: 2-3 weeks
**Risk**: LOW (existing code keeps working, new code available)

### Phase 3: MCP Tool Migration

**Goal**: Update MCP tool to use new API + new action names

**Changes**:
1. ✅ Add new actions: `create`, `update-mcode`, `update-and-refresh`, `load-to`, `unload`, `refresh-all`
2. ✅ Keep old actions as aliases (backwards compatibility)
3. ✅ Update prompts with new workflow guidance
4. ✅ Add deprecation warnings for old actions
5. ✅ Update documentation

**Timeline**: 1-2 weeks
**Risk**: MEDIUM (LLMs may use cached prompts, need clear migration path)

### Phase 4: Breaking Change Release

**Goal**: Remove old methods, finalize API

**Changes**:
1. ✅ Remove `[Obsolete]` methods: `ImportAsync`, `SetLoad*Async`
2. ✅ Remove old MCP action aliases
3. ✅ Major version bump (v2.0.0)
4. ✅ Publish migration guide

**Timeline**: 1 week
**Risk**: HIGH (breaking changes, user migration required)

---

## Testing Strategy

### Diagnostic Tests (Preserve & Expand)

**Current Tests** (from `ExcelQueryTableBehaviorDiagnostics.Split.cs`):
- ✅ Test 1: Load to worksheet creates single QueryTable
- ✅ Test 2: Multiple refreshes don't create duplicates
- ✅ Test 3: M code updates work in isolation
- ✅ Test 4: Connection-only has no QueryTable
- ✅ Test 5: Loading connection-only creates QueryTable

**Additional Tests Needed** (Future API validation):
- ✅ Test 6: `UpdateMCodeAsync()` doesn't trigger refresh
- ✅ Test 7: `UpdateAndRefreshAsync()` updates + refreshes atomically
- ✅ Test 8: `LoadToAsync()` is atomic (config + refresh in one call)
- ✅ Test 9: `UnloadAsync()` removes QueryTable safely
- ✅ Test 10: `RefreshAllAsync()` refreshes all queries efficiently

### Integration Tests (MCP Layer)

**Test Scenarios**:
1. ✅ Create workflow: Single `create()` call validates + loads data
2. ✅ Update workflow: `update-mcode()` doesn't refresh, `refresh()` does
3. ✅ Load destination changes: `load-to()` is atomic
4. ✅ Batch updates: `update-mcode()` × N + `refresh-all()` efficient
5. ✅ Error handling: Excel stress scenarios handled gracefully

---

## Migration Guide (for LLMs & Users)

### Quick Reference

| **Old Pattern** | **New Pattern** | **Why Change** |
|-----------------|-----------------|----------------|
| `import(loadDestination='worksheet')` | `create(loadTo='Worksheet')` | Clearer verb, explicit enum |
| `import(loadDestination='connection-only')` + `refresh()` + `set-load-to-table()` | `create(loadTo='Worksheet')` | Single atomic call, no double refresh |
| `update()` | `update-mcode()` or `update-and-refresh()` | Explicit refresh control |
| `set-load-to-table()` | `load-to(destination='Worksheet')` | Clear intent, consistent naming |
| `set-load-to-data-model()` | `load-to(destination='DataModel')` | Clear intent, consistent naming |
| `set-load-to-both()` | `load-to(destination='Both')` | Clear intent, consistent naming |
| `set-connection-only()` | `unload()` | Clearer action name |

### Code Examples

**Example 1: Create & Load**
```csharp
// OLD
await commands.ImportAsync(batch, "Sales", "sales.pq", loadDestination: "worksheet", worksheetName: "SalesData");

// NEW
await commands.CreateAsync(batch, "Sales", "sales.pq", QueryDestination.Worksheet, "SalesData");
```

**Example 2: Update M Code Without Refresh**
```csharp
// OLD (unclear if refresh happens)
await commands.UpdateAsync(batch, "Sales", "sales-v2.pq");

// NEW (explicit: no refresh)
await commands.UpdateMCodeAsync(batch, "Sales", "sales-v2.pq");
```

**Example 3: Update M Code With Refresh**
```csharp
// OLD (manual 2-step)
await commands.UpdateAsync(batch, "Sales", "sales-v2.pq");
await commands.RefreshAsync(batch, "Sales");

// NEW (atomic operation)
await commands.UpdateAndRefreshAsync(batch, "Sales", "sales-v2.pq");
```

**Example 4: Load to Data Model**
```csharp
// OLD
await commands.SetLoadToDataModelAsync(batch, "Sales");

// NEW
await commands.LoadToAsync(batch, "Sales", QueryDestination.DataModel);
```

**Example 5: Batch Update Queries**
```csharp
// OLD (each update may trigger refresh)
await commands.UpdateAsync(batch, "Query1", "code1.pq");
await commands.UpdateAsync(batch, "Query2", "code2.pq");
await commands.UpdateAsync(batch, "Query3", "code3.pq");

// NEW (explicit batch control)
await commands.UpdateMCodeAsync(batch, "Query1", "code1.pq");
await commands.UpdateMCodeAsync(batch, "Query2", "code2.pq");
await commands.UpdateMCodeAsync(batch, "Query3", "code3.pq");
await commands.RefreshAllAsync(batch);  // Single batch refresh
```

---

## Success Metrics

### Performance Improvements

| **Workflow** | **Current** | **Future** | **Improvement** |
|--------------|-------------|------------|-----------------|
| Create + load to worksheet | 1 call, 1 refresh | 1 call, 1 refresh | ✅ Same (already optimal) |
| Create connection-only + load later | 3 calls, 2 refreshes | 1 call, 1 refresh | ✅ 67% fewer operations |
| Batch update 5 queries | 5-10 operations | 6 operations (5 updates + 1 refresh) | ✅ 40-50% fewer operations |
| Change load destination | 1 call, 1 refresh | 1 call, 1 refresh | ✅ Same (already atomic) |

### LLM Usability Improvements

| **Metric** | **Current** | **Future** | **Improvement** |
|------------|-------------|------------|-----------------|
| Action name clarity | ⚠️ "set-load-to-X" confusing | ✅ "load-to", "unload" clear | Cognitive load reduced |
| Parameter consistency | ⚠️ `loadDestination: string` | ✅ `loadTo: enum` | Type safety, autocomplete |
| Workflow efficiency | ⚠️ 3 steps for simple tasks | ✅ 1 step atomic operations | LLM trial-error eliminated |
| Error messages | ⚠️ Generic COM errors | ✅ Structured errors with IsRetryable | LLMs handle retries |
| Error resilience | ⚠️ RPC timeout kills operation | ✅ Clear error metadata for LLM retry logic | Better UX under stress |

### Code Quality Improvements

| **Metric** | **Current** | **Future** | **Improvement** |
|------------|-------------|------------|-----------------|
| Defensive cleanup code | ✅ Present (unnecessary) | ✅ Removed | Code size reduced |
| Refresh control granularity | ⚠️ Implicit refresh behavior | ✅ Explicit refresh methods | Predictable behavior |
| Error metadata | ⚠️ Generic COM errors | ✅ IsRetryable, ErrorCategory, SuggestedActions | LLMs orchestrate retries |

---

## Appendix A: Diagnostic Test Evidence

### Test Results Summary

| Test | Duration | QueryTables Count | Outcome | Key Finding |
|------|----------|-------------------|---------|-------------|
| 1 - Load to Worksheet | 17s | 1 | ✅ PASS | Single QueryTable created |
| 2 - Refresh 4 Times | 18s | 1 (all refreshes) | ✅ PASS | **NO duplicates created** |
| 3 - Update M Code | 18s | 1 | ✅ PASS | **Refresh succeeded in isolation** |
| 4 - Connection-Only | 12s | 0 | ✅ PASS | **NO auto-QueryTable** |
| 5 - Load Connection-Only | 16s | 1 | ✅ PASS | Manual QueryTable works |

### Code Impact from Diagnostics

**Confirmed Safe Removals**:
```csharp
// ❌ REMOVE: Unnecessary defensive code
while (queryTables.Count > 1) {
    queryTables.Item(queryTables.Count).Delete();
}
```

**Required Error Handling**:
```csharp
// ✅ ADD: Return structured error info for LLM retry logic
catch (COMException ex) when (ex.HResult == unchecked((int)0x800706BE)) {
    return new OperationResult
    {
        Success = false,
        ErrorMessage = "Excel COM timeout (0x800706BE). Operation may succeed if retried.",
        ErrorCategory = "ExcelStress",
        IsRetryable = true,
        RetryGuidance = "LLM can implement exponential backoff (e.g., wait 1s, 2s, 4s)",
        SuggestedNextActions = new[]
        {
            "Retry operation (LLMs handle retry orchestration)",
            "Alternative: Close and reopen workbook after max retries"
        }
    };
}
```

**LLM Implementation Example**:
```python
# LLM can implement retry logic based on IsRetryable flag
async def update_query_with_retry(query_name, code_file, max_retries=3):
    for attempt in range(max_retries):
        result = await excel_powerquery(
            action='update-mcode',
            queryName=query_name,
            sourcePath=code_file
        )
        
        if result['success']:
            return result
            
        # Check if retryable
        if not result.get('isRetryable', False):
            raise Exception(f"Non-retryable error: {result['errorMessage']}")
        
        # Exponential backoff
        wait_time = 2 ** attempt  # 1s, 2s, 4s
        await asyncio.sleep(wait_time)
    
    raise Exception(f"Max retries exceeded for {query_name}")
```

---

## Appendix B: Current vs Future API Complete Mapping

### Interface Comparison

```csharp
// ============================================
// CURRENT API (IPowerQueryCommands.cs)
// ============================================

// CRUD (6 methods)
Task<PowerQueryListResult> ListAsync(IExcelBatch batch);
Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName);
Task<OperationResult> ImportAsync(IExcelBatch batch, string queryName, string mCodeFile, string loadDestination = "worksheet", string? worksheetName = null);
Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile);
Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName);

// Refresh (2 methods)
Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName);
Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName, TimeSpan? timeout);

// Load Config (5 methods)
Task<OperationResult> SetLoadToTableAsync(IExcelBatch batch, string queryName, string sheetName);
Task<OperationResult> SetLoadToDataModelAsync(IExcelBatch batch, string queryName);
Task<OperationResult> SetLoadToBothAsync(IExcelBatch batch, string queryName, string sheetName);
Task<OperationResult> SetConnectionOnlyAsync(IExcelBatch batch, string queryName);
Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName);

// Advanced (5 methods)
Task<PowerQueryViewResult> ErrorsAsync(IExcelBatch batch, string queryName);
Task<OperationResult> LoadToAsync(IExcelBatch batch, string queryName, string sheetName);
Task<WorksheetListResult> ListExcelSourcesAsync(IExcelBatch batch);
Task<PowerQueryViewResult> EvalAsync(IExcelBatch batch, string mExpression);

// TOTAL: 18 methods

// ============================================
// FUTURE API (Proposed)
// ============================================

// CRUD (6 methods)
Task<PowerQueryListResult> ListAsync(IExcelBatch batch);
Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName);
Task<PowerQueryCreateResult> CreateAsync(IExcelBatch batch, string queryName, string mCodeFile, QueryDestination loadTo = QueryDestination.Worksheet, string? worksheetName = null);
Task<OperationResult> UpdateMCodeAsync(IExcelBatch batch, string queryName, string mCodeFile);
Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName);

// Data Load (4 methods)
Task<PowerQueryLoadResult> LoadToAsync(IExcelBatch batch, string queryName, QueryDestination destination, string? worksheetName = null);
Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName, TimeSpan? timeout = null);
Task<OperationResult> UnloadAsync(IExcelBatch batch, string queryName);
Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName);

// Advanced (4 methods)
Task<WorksheetListResult> ListExcelSourcesAsync(IExcelBatch batch);
Task<PowerQueryEvalResult> ValidateSyntaxAsync(IExcelBatch batch, string mExpression);
Task<PowerQueryRefreshResult> UpdateAndRefreshAsync(IExcelBatch batch, string queryName, string mCodeFile, TimeSpan? timeout = null);
Task<PowerQueryRefreshAllResult> RefreshAllAsync(IExcelBatch batch, TimeSpan? timeout = null);

// TOTAL: 14 methods (-4 from current, unified & simplified)
// REMOVED: ErrorsAsync (Excel COM API limitation - no error properties available)
// RENAMED: EvalAsync → ValidateSyntaxAsync (clarifies limited scope)
```

---

## Conclusion

This specification proposes a **comprehensive redesign** of the Power Query API that:

1. ✅ **Eliminates inefficient workflows** discovered through LLM usage analysis
2. ✅ **Simplifies API surface** from 18 methods to 14 methods (-22%)
3. ✅ **Improves LLM discoverability** with clear action names matching user intent
4. ✅ **Integrates diagnostic findings** to remove unnecessary defensive code
5. ✅ **Provides granular control** for power users (update M code without refresh)
6. ✅ **Maintains backwards compatibility** during transition period
7. ✅ **Improves performance** by eliminating double-refresh patterns

**Next Steps**:
1. Review & approve specification
2. Implement Phase 1 (internal refactoring, no breaking changes)
3. Implement Phase 2 (add new methods alongside old ones)
4. Implement Phase 3 (MCP tool migration with aliases)
5. Release Phase 4 (breaking changes in v2.0.0)

**Timeline**: 5-7 weeks for complete migration
**Risk**: Managed through phased rollout with deprecation period

---

**Document Version**: 1.0  
**Last Updated**: 2025-01-29  
**Status**: PROPOSED - Awaiting Review

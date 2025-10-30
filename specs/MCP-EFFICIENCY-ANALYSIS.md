# MCP Server Efficiency Analysis - Three Critical API Issues

**Date**: October 30, 2025  
**Reported Issue**: LLM workflow forced to make 2× calls for Data Model queries + sensitivity label failures  
**Source**: Real-world usage report from GitHub Copilot  
**Severity**: CRITICAL - API design flaws + Excel privacy bug + inefficient bulk operations

---

## TL;DR - Three Critical Issues Identified

**Date**: October 30, 2025  
**Analyzed by**: GitHub Copilot (AI Assistant analyzing another AI's workflow)  
**Source**: Real-world bug report from LLM using ExcelMcp MCP Server

### Issue 1: Missing loadDestination Parameter (API Design Flaw)

**Problem**: `import` action can load to worksheet but NOT to Data Model in one call.

**Current forced pattern**:
```typescript
excel_powerquery({ action: "import", loadToWorksheet: false })       // Step 1: Import M code
excel_powerquery({ action: "set-load-to-data-model" })                // Step 2: Load to Data Model
```

**What LLMs need**:
```typescript
excel_powerquery({ 
  action: "import", 
  loadDestination: "data-model"  // ✅ One operation, clear intent
})
```

**Impact**: 
- LLM made **8 calls instead of 4** for Data Model workflow
- **50% wasted calls** due to API limitation
- **Breaking change acceptable** - API is designed for LLMs, not humans

---

### Issue 2: File.Contents() Sensitivity Label Bug (NOT a Privacy Issue!)

**Problem**: Queries using `File.Contents()` to read external Excel files fail when loading to Data Model if source file has Microsoft Purview sensitivity labels:

```
[DataSource.Error] We can't load data from this source because it's protected. 
You may not have permissions to apply its required sensitivity label...
```

**Root Cause**: **Power Query cannot access Excel files with sensitivity labels** (other than "Public" or "Non-Business") because labeled files are encrypted.

Microsoft docs: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations

> "Power Query Online is unable to access encrypted Excel files. Since Excel files labeled with sensitivity types other than 'Public' or 'Non-Business' are encrypted, they aren't accessible through Power Query Online."

**Current behavior**:
- ✅ Regions query: SUCCESS (sensitivity label was removed from file)
- ❌ Milestones query: FAILED - source file still has sensitivity label
- **File.Contents() + Excel.Workbook() fails with encrypted files**

**Manual workaround LLM discovered** (doesn't solve root problem):
```typescript
excel_powerquery({ action: "set-load-to-table" })    // Still fails - can't read encrypted file!
excel_table({ action: "create" })                     
excel_table({ action: "add-to-datamodel" })           
```

**CORRECT Solution**: Guide LLM to fix the M code or source file

When `set-load-to-data-model` fails with this error pattern:
1. **Detect** error contains "protected" or "sensitivity label"
2. **Parse M code** to extract file path from `File.Contents()`
3. **Provide actionable guidance**:

```
Error: Source Excel file has Microsoft Purview sensitivity labels (file encryption).

Power Query cannot access encrypted Excel files.

SOLUTION: Change sensitivity label to Public
  - Open: [extracted file path]
  - Click Home tab → Sensitivity button → Select "Public" label
  - Save and close
  - Retry: excel_powerquery({ action: "set-load-to-data-model", queryName: "{queryName}" })

For details: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations
```

**Implementation**:
```csharp
// In PowerQueryCommands.SetLoadToDataModelAsync
catch (COMException ex) when (ex.Message.Contains("protected") || ex.Message.Contains("sensitivity label"))
{
    var mCode = await GetQueryMCodeAsync(batch, queryName);
    var filePath = ExtractFileContentsPath(mCode); // Parse "File.Contents("path")"
    
    return new OperationResult
    {
        Success = false,
        ErrorMessage = $@"Source Excel file has Microsoft Purview sensitivity labels (encryption).

Power Query cannot read encrypted Excel files.

SOLUTION: Change sensitivity label to Public
  File: {filePath}
  Steps: Home tab → Sensitivity button → Select ""Public"" label → Save
  Then: Retry set-load-to-data-model

Details: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations"
    };
}
```

**Why NOT automatic fallback**:
- ❌ Fallback to worksheet → table → Data Model **doesn't solve the problem**
- ❌ Power Query still can't read the encrypted source file
- ❌ Hides the real issue from the LLM
- ✅ LLM needs to know about this Power Query limitation
- ✅ User must either remove label OR change M code

**Impact**:
- ✅ **Educates LLM** about Microsoft Purview limitation
- ✅ **Actionable guidance** with file path and options
- ✅ **Root cause fix** - resolves actual encryption problem
- ✅ **No wasted MCP calls** - LLM knows what to do immediately

**Effort**: ~2 hours (error detection, M code parsing for file path, helpful error message)

---

### Issue 3: No Bulk Parameter Creation (10× Excel Sessions)

**Problem**: Creating 5 parameters = 10 MCP calls = 10 Excel sessions

**Current pattern**:
```typescript
excel_parameter({ action: "create", parameterName: "Start_Date", value: "Sheet1!$A$1" })
excel_parameter({ action: "set", parameterName: "Start_Date", value: "2025-07-01" })
// ... 8 more calls
```

**What LLMs need**:
```typescript
excel_parameter({
  action: "create-bulk",
  parameters: [
    { name: "Start_Date", reference: "Sheet1!$A$1", value: "2025-07-01" },
    { name: "Duration_Months", reference: "Sheet1!$B$1", value: 12 },
    // ... 3 more parameters
  ]
})
```

**Impact**:
- **90% reduction** in parameter creation calls (10 → 1)
- **10× fewer Excel sessions** for parameter setup

---

### Combined Impact

**Current workflow**: 24 MCP calls, 23 Excel sessions, ~46-69 seconds  
**Optimized workflow**: 11 MCP calls, 3 Excel sessions, ~9-12 seconds  
**Improvement**: **54% fewer calls, 87% fewer Excel sessions, 80% faster**

**Recommendations**:
1. **Add `loadDestination` parameter** to import action (BREAKING CHANGE - acceptable)
2. **Helpful error message** for sensitivity label issues (NO BREAKING CHANGE - user education)
3. **Add `create-bulk` action** for parameters (NO BREAKING CHANGE)

**Total effort**: ~11 hours to implement all three fixes

---

## The Missing Parameter

### Current API (Incomplete)

```csharp
public static async Task<string> ExcelPowerQuery(
    string action,
    string excelPath,
    string? queryName = null,
    string? sourcePath = null,
    bool? loadToWorksheet = null,  // ✅ Supports worksheet
    // ❌ MISSING: bool? loadToDataModel = null
)
```

**What LLM can do**:
- ✅ Import and load to worksheet: `loadToWorksheet: true`
- ✅ Import as connection-only: `loadToWorksheet: false`
- ❌ **Import and load to Data Model**: NOT POSSIBLE

### What Should Exist

```csharp
public static async Task<string> ExcelPowerQuery(
    string action,
    string excelPath,
    string? queryName = null,
    string? sourcePath = null,
    bool? loadToWorksheet = null,
    bool? loadToDataModel = null,  // ✅ ADD THIS
    bool? loadToBoth = null         // ✅ OR THIS (worksheet + Data Model)
)
```

**Behavior matrix**:
| loadToWorksheet | loadToDataModel | Result |
|----------------|-----------------|---------|
| true | false | Load to worksheet only (current default) |
| false | true | Load to Data Model only (**NEW**) |
| true | true | Load to both (**NEW**) |
| false | false | Connection-only (current) |

---

---

## Real-World Impact Analysis

### Actual LLM Workflow (From Bug Report)

```typescript
// Step 1-4: Import 4 queries as connection-only (4 MCP calls, 4 Excel sessions)
excel_powerquery({ action: "import", queryName: "ProjectRootDirectory Parameter", sourcePath: "...", loadToWorksheet: false })
excel_powerquery({ action: "import", queryName: "Parameters", sourcePath: "...", loadToWorksheet: false })
excel_powerquery({ action: "import", queryName: "Regions", sourcePath: "...", loadToWorksheet: false })
excel_powerquery({ action: "import", queryName: "Milestones", sourcePath: "...", loadToWorksheet: false })

// Step 5-8: Load each to Data Model separately (4 MCP calls, 4 Excel sessions)
excel_powerquery({ action: "set-load-to-data-model", queryName: "ProjectRootDirectory Parameter" })
excel_powerquery({ action: "set-load-to-data-model", queryName: "Parameters" })
excel_powerquery({ action: "set-load-to-data-model", queryName: "Regions" })
excel_powerquery({ action: "set-load-to-data-model", queryName: "Milestones" })  // ❌ FAILS

// TOTAL: 8 calls, 8 Excel sessions (one fails)
```

**Why LLM does this**: 
- API doesn't offer `loadToDataModel` parameter on import
- LLM correctly sets `loadToWorksheet: false` to avoid loading to worksheet
- Then LLM is forced to make separate call to configure Data Model loading

---

### What LLM SHOULD Be Able To Do

```typescript
// Option A: With new loadToDataModel parameter (4 MCP calls)
excel_powerquery({ action: "import", queryName: "ProjectRootDirectory Parameter", sourcePath: "...", loadToDataModel: true })
excel_powerquery({ action: "import", queryName: "Parameters", sourcePath: "...", loadToDataModel: true })
excel_powerquery({ action: "import", queryName: "Regions", sourcePath: "...", loadToDataModel: true })
excel_powerquery({ action: "import", queryName: "Milestones", sourcePath: "...", loadToDataModel: true })

// TOTAL: 4 calls instead of 8 (50% reduction)

// Option B: With batch mode + new parameter (1 Excel session)
batch = begin_excel_batch({ excelPath })
excel_powerquery({ action: "import", queryName: "Q1", sourcePath: "...", loadToDataModel: true, batchId })
excel_powerquery({ action: "import", queryName: "Q2", sourcePath: "...", loadToDataModel: true, batchId })
excel_powerquery({ action: "import", queryName: "Q3", sourcePath: "...", loadToDataModel: true, batchId })
excel_powerquery({ action: "import", queryName: "Q4", sourcePath: "...", loadToDataModel: true, batchId })
commit_excel_batch({ batchId, save: true })

// TOTAL: 6 calls, 1 Excel session (massive improvement)
```

---

```
CURRENT LLM WORKFLOW (23 Excel sessions):
┌─────────────────────────────────────────────────────────────────┐
│ Create File     [Excel Launch #1 ]  ~2s                        │
│ Param Create    [Excel Launch #2 ]  ~2s                        │
│ Param Set       [Excel Launch #3 ]  ~2s                        │
│ Param Create    [Excel Launch #4 ]  ~2s                        │
│ Param Set       [Excel Launch #5 ]  ~2s                        │
│ ... (6 more param operations)       ~12s                        │
│ Import Query 1  [Excel Launch #12]  ~2s                        │
│ Import Query 2  [Excel Launch #13]  ~2s                        │
│ Import Query 3  [Excel Launch #14]  ~2s                        │
│ Import Query 4  [Excel Launch #15]  ~2s                        │
│ Load to DM Q1   [Excel Launch #16]  ~3s                        │
│ Load to DM Q2   [Excel Launch #17]  ~3s                        │
│ Load to DM Q3   [Excel Launch #18]  ~3s                        │
│ Load to DM Q4   [Excel Launch #19]  ~3s  ❌ FAILS              │
│ Workaround 1    [Excel Launch #20]  ~3s                        │
│ Workaround 2    [Excel Launch #21]  ~2s                        │
│ Workaround 3    [Excel Launch #22]  ~2s                        │
│ Verify          [Excel Launch #23]  ~2s                        │
│                                                                 │
│ TOTAL: 23 Excel sessions = ~46-69 seconds                      │
└─────────────────────────────────────────────────────────────────┘

OPTIMAL WORKFLOW WITH BATCH MODE (2 sessions):
┌─────────────────────────────────────────────────────────────────┐
│ [Excel Launch #1 - BATCH SESSION]                              │
│ ├─ Create File                      ~0.1s                       │
│ ├─ Set Values (bulk)                ~0.1s                       │
│ ├─ Create Param 1-5                 ~0.5s                       │
│ ├─ Import Query 1-4                 ~0.8s                       │
│ ├─ Load to DM Q1-3                  ~2.0s                       │
│ ├─ Workaround Q4 (3 operations)     ~1.0s                       │
│ └─ Save & Close                     ~0.5s                       │
│                          Subtotal: ~5 seconds                   │
│                                                                 │
│ [Excel Launch #2 - VERIFY]                                     │
│ └─ Verify Data Model                ~3s                         │
│                                                                 │
│ TOTAL: 2 Excel sessions = ~8 seconds                           │
│ SAVINGS: 38-61 seconds (85% faster!)                           │
└─────────────────────────────────────────────────────────────────┘
```

---

---

## Root Cause: API Design Flaw

### Core ImportAsync Method (Limited)

```csharp
// File: src/ExcelMcp.Core/Commands/PowerQueryCommands.cs
public async Task<OperationResult> ImportAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile, 
    bool loadToWorksheet = true,  // ✅ Only worksheet supported
    string? worksheetName = null)
{
    // Imports M code and optionally loads to worksheet
    // ❌ NO DATA MODEL SUPPORT
}
```

**What's missing**: No `loadToDataModel` parameter or logic

---

### MCP Tool Parameters (Incomplete)

```csharp
// File: src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs
[Description("Automatically load query data to worksheet for validation (default: true). 
             When false, creates connection-only query without validation.")]
bool? loadToWorksheet = null,

// ❌ MISSING:
// [Description("Load query data to Power Pivot Data Model (default: false)")]
// bool? loadToDataModel = null,
```

**Result**: LLM has NO WAY to import and load to Data Model in one operation

---

### Why This Matters

**Current user intent**: "Import these 4 .pq files into the Data Model"

**What LLM must do**:
1. Import query (don't load anywhere)
2. Configure to load to Data Model

**What LLM should do**:
1. Import query AND load to Data Model (one operation)

**Developer surprise**: 
- You can `import` + load to worksheet in one call
- You CANNOT `import` + load to Data Model in one call
- Inconsistent API design

---

## Issue 2: No Batch Mode Usage (Secondary)

### Issue 1: No Batch Mode for Data Model Loading (4× slower)
- ❌ Current: 4 separate `set-load-to-data-model` calls = **4 Excel sessions**
- ✅ Efficient: 1 batch session with 4 operations = **1 Excel session**
- 💰 **Performance impact**: ~8-12 seconds wasted (4 Excel startups vs 1)

### Issue 2: Inefficient Parameter Creation (40% wasted calls)
- ❌ Current: **10 MCP calls** (create + set for each of 5 parameters)
- ✅ Efficient: **6 calls** (bulk value write + 5 creates)
- 💰 **Performance impact**: ~4 seconds wasted

### Issue 3: No Batch Mode for Parameter Creation (10× slower)
- ❌ Current: 10 separate calls = **10 Excel sessions**
- ✅ Efficient with batch: 1 batch session with 6 operations = **1 Excel session**
- 💰 **Performance impact**: ~18-20 seconds wasted

**Total waste in this workflow**: ~30-36 seconds of unnecessary Excel session overhead

This inefficiency stems from:

1. **LLM not aware of batch mode**: Documentation mentions it, but LLM doesn't use it
2. **API design confusion**: The parameter API splits creation into two separate concerns
3. **Missing best practices**: No clear guidance on when/how to use batch mode

---

## Issue 1: Data Model Loading Without Batch Mode (CRITICAL)

### Current Pattern (4 Excel sessions - extremely inefficient)

```typescript
// Query 1: Opens Excel, loads query, saves, closes Excel
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "file.xlsx",
  queryName: "ProjectRootDirectory Parameter" 
})

// Query 2: Opens Excel AGAIN, loads query, saves, closes Excel
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "file.xlsx",
  queryName: "Parameters" 
})

// Query 3: Opens Excel AGAIN, loads query, saves, closes Excel
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "file.xlsx",
  queryName: "Regions" 
})

// Query 4: Opens Excel AGAIN, loads query, saves, closes Excel
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "file.xlsx",
  queryName: "Milestones" 
})

// Total: 4 Excel launches, 4 file opens, 4 saves, 4 closes
// Time: ~12-16 seconds (3-4 seconds per operation)
```

**Why the LLM does this**:
- Each call mentions batch mode in the hint: "For configuring multiple queries, use begin_excel_batch"
- But LLM apparently ignores this hint or doesn't understand the pattern
- No clear example showing HOW to use batch mode for this workflow

---

### Efficient Pattern (1 Excel session)

```typescript
// Step 1: Begin batch session (opens Excel once)
const batch = begin_excel_batch({ 
  excelPath: "file.xlsx" 
})

// Step 2: Configure all 4 queries in SAME Excel session
excel_powerquery({ 
  action: "set-load-to-data-model", 
  queryName: "ProjectRootDirectory Parameter",
  batchId: batch.batchId 
})

excel_powerquery({ 
  action: "set-load-to-data-model", 
  queryName: "Parameters",
  batchId: batch.batchId 
})

excel_powerquery({ 
  action: "set-load-to-data-model", 
  queryName: "Regions",
  batchId: batch.batchId 
})

excel_powerquery({ 
  action: "set-load-to-data-model", 
  queryName: "Milestones",
  batchId: batch.batchId 
})

// Step 3: Save and close Excel ONCE
commit_excel_batch({ 
  batchId: batch.batchId, 
  save: true 
})

// Total: 1 Excel launch, 1 file open, 1 save, 1 close
// Time: ~3-4 seconds total
// Savings: ~8-12 seconds (75% faster)
```

**Benefits**:
- ✅ **75% faster** (4 seconds vs 16 seconds)
- ✅ **Atomic operation** - all 4 queries loaded or none
- ✅ **More reliable** - single transaction, fewer points of failure
- ✅ **Less Excel churn** - file opened/closed once instead of 4 times

---

## Issue 2: Parameter Creation Without Batch Mode

### Current Pattern (10 Excel sessions - extremely inefficient)

```typescript
// Step 1: Create named range pointing to Sheet1!$A$1 (Excel session 1)
excel_parameter({ action: "create", parameterName: "Start_Date", value: "Sheet1!$A$1" })

// Step 2: Set value in cell A1 (Excel session 2)
excel_parameter({ action: "set", parameterName: "Start_Date", value: "2025-07-01" })

// Step 3: Create named range pointing to Sheet1!$B$1 (Excel session 3)
excel_parameter({ action: "create", parameterName: "Duration_Months", value: "Sheet1!$B$1" })

// Step 4: Set value in cell B1 (Excel session 4)
excel_parameter({ action: "set", parameterName: "Duration_Months", value: "36" })

// ... repeat for 5 parameters
// Total: 10 Excel sessions (10 launches, 10 opens, 10 saves, 10 closes)
// Time: ~20-30 seconds
```

**Why the LLM does this**:
- API requires cell reference for `create` action
- Separate `set` action required to populate cell value
- No indication that there's a more efficient way

---

### Efficient Pattern Option A: Better API usage, no batch (6 Excel sessions)

```typescript
// Step 1: Populate all parameter values in ONE call (Excel session 1)
excel_range({ 
  action: "set-values", 
  sheetName: "Sheet1", 
  rangeAddress: "A1:E1",
  values: [["2025-07-01", "36", "FY26 Consumption Plan", "PhysicsX", "D:\\source\\repos\\cp_toolkit"]]
})

// Step 2-6: Create named ranges (Excel sessions 2-6)
excel_parameter({ action: "create", parameterName: "Start_Date", value: "Sheet1!$A$1" })
excel_parameter({ action: "create", parameterName: "Duration_Months", value: "Sheet1!$B$1" })
excel_parameter({ action: "create", parameterName: "Plan_Name", value: "Sheet1!$C$1" })
excel_parameter({ action: "create", parameterName: "Customer_Name", value: "Sheet1!$D$1" })
excel_parameter({ action: "create", parameterName: "ProjectRoot", value: "Sheet1!$E$1" })

// Total: 6 Excel sessions
// Time: ~12-18 seconds
// Savings: ~8-12 seconds vs current (40% faster)
```

---

### Efficient Pattern Option B: WITH batch mode (1 Excel session - BEST)

```typescript
// Step 1: Begin batch session
const batch = begin_excel_batch({ excelPath: "file.xlsx" })

// Step 2: Populate all values in batch
excel_range({ 
  action: "set-values", 
  sheetName: "Sheet1", 
  rangeAddress: "A1:E1",
  values: [["2025-07-01", "36", "FY26 Consumption Plan", "PhysicsX", "D:\\source\\repos\\cp_toolkit"]],
  batchId: batch.batchId
})

// Step 3: Create all named ranges in batch
excel_parameter({ action: "create", parameterName: "Start_Date", value: "Sheet1!$A$1", batchId: batch.batchId })
excel_parameter({ action: "create", parameterName: "Duration_Months", value: "Sheet1!$B$1", batchId: batch.batchId })
excel_parameter({ action: "create", parameterName: "Plan_Name", value: "Sheet1!$C$1", batchId: batch.batchId })
excel_parameter({ action: "create", parameterName: "Customer_Name", value: "Sheet1!$D$1", batchId: batch.batchId })
excel_parameter({ action: "create", parameterName: "ProjectRoot", value: "Sheet1!$E$1", batchId: batch.batchId })

// Step 4: Save and close
commit_excel_batch({ batchId: batch.batchId, save: true })

// Total: 1 Excel session (7 operations batched)
// Time: ~2-3 seconds
// Savings: ~17-27 seconds vs current (90% faster!)
```

```typescript
// Step 1: Populate all parameter values in ONE call
excel_range({ 
  action: "set-values", 
  sheetName: "Sheet1", 
  rangeAddress: "A1:E1",
  values: [["2025-07-01", "36", "FY26 Consumption Plan", "PhysicsX", "D:\\source\\repos\\cp_toolkit"]]
})

// Step 2-6: Create named ranges (5 calls)
excel_parameter({ action: "create", parameterName: "Start_Date", value: "Sheet1!$A$1" })
excel_parameter({ action: "create", parameterName: "Duration_Months", value: "Sheet1!$B$1" })
excel_parameter({ action: "create", parameterName: "Plan_Name", value: "Sheet1!$C$1" })
excel_parameter({ action: "create", parameterName: "Customer_Name", value: "Sheet1!$D$1" })
excel_parameter({ action: "create", parameterName: "ProjectRoot", value: "Sheet1!$E$1" })

// Total: 6 calls (40% reduction)
```

---

## Combined Workflow: Full Efficiency Analysis

Let's analyze the **complete workflow** from the bug report:

### Phase 1: Current LLM Workflow (Actual from report)

```typescript
// Step 1: Create workbook (1 session)
excel_file({ action: "create-empty", excelPath: "..." })

// Step 2-11: Create 5 parameters without batch (10 sessions)
excel_parameter({ action: "create", ... })  // × 5
excel_parameter({ action: "set", ... })     // × 5

// Step 12-15: Import 4 queries (4 sessions)
excel_powerquery({ action: "import", loadToWorksheet: false, ... })  // × 4

// Step 16-19: Load to Data Model without batch (4 sessions)
excel_powerquery({ action: "set-load-to-data-model", ... })  // × 4

// Step 20-23: Workaround for Milestones (3 sessions)
excel_powerquery({ action: "set-load-to-table", ... })
excel_table({ action: "create", ... })
excel_table({ action: "add-to-datamodel", ... })

// Step 24: Verify (1 session)
excel_datamodel({ action: "list-tables", ... })

// TOTAL: 24 MCP calls, 23 Excel sessions, ~46-69 seconds
```

---

### Phase 2: Optimized Workflow (Better API usage, still no batch)

```typescript
// Step 1: Create workbook (1 session)
excel_file({ action: "create-empty", excelPath: "..." })

// Step 2: Populate all parameter cells (1 session)
excel_range({ action: "set-values", sheetName: "Sheet1", rangeAddress: "A1:E1", values: [[...]] })

// Step 3-7: Create 5 named ranges (5 sessions)
excel_parameter({ action: "create", ... })  // × 5

// Step 8-11: Import 4 queries (4 sessions)
excel_powerquery({ action: "import", loadToWorksheet: false, ... })  // × 4

// Step 12-15: Load to Data Model (4 sessions)
excel_powerquery({ action: "set-load-to-data-model", ... })  // × 4

// Step 16-18: Workaround (3 sessions)
excel_powerquery({ action: "set-load-to-table", ... })
excel_table({ action: "create", ... })
excel_table({ action: "add-to-datamodel", ... })

// Step 19: Verify (1 session)
excel_datamodel({ action: "list-tables", ... })

// TOTAL: 19 MCP calls, 19 Excel sessions, ~38-57 seconds
// SAVINGS: 5 calls, 4 sessions, ~8-12 seconds (20% faster)
```

---

### Phase 3: OPTIMAL Workflow (WITH batch mode - BEST PRACTICE)

```typescript
// Step 1: Begin batch session
const batch = begin_excel_batch({ excelPath: "file.xlsx" })

// Step 2: Create workbook in batch
excel_file({ action: "create-empty", excelPath: "...", batchId: batch.batchId })

// Step 3: Populate parameter cells in batch
excel_range({ 
  action: "set-values", 
  sheetName: "Sheet1", 
  rangeAddress: "A1:E1", 
  values: [[...]],
  batchId: batch.batchId 
})

// Step 4-8: Create named ranges in batch
excel_parameter({ action: "create", ..., batchId: batch.batchId })  // × 5

// Step 9-12: Import queries in batch
excel_powerquery({ action: "import", loadToWorksheet: false, ..., batchId: batch.batchId })  // × 4

// Step 13-16: Load to Data Model in batch
excel_powerquery({ action: "set-load-to-data-model", ..., batchId: batch.batchId })  // × 4
// (OR use workaround if File.Contents detected)

// Step 17: Save and close ONCE
commit_excel_batch({ batchId: batch.batchId, save: true })

// Step 18: Verify in new session
excel_datamodel({ action: "list-tables", ... })

// TOTAL: 19 MCP calls, 2 Excel sessions, ~8-12 seconds
// SAVINGS: 5 calls, 21 sessions, ~38-57 seconds (85% faster!)
```

---

## Performance Comparison Table

| Workflow | MCP Calls | Excel Sessions | Estimated Time | Savings |
|----------|-----------|----------------|----------------|---------|
| **Current (Actual)** | 24 | 23 | 46-69 seconds | Baseline |
| **Better API** | 19 | 19 | 38-57 seconds | 8-12s (20%) |
| **WITH Batch** | 19 | 2 | 8-12 seconds | **38-57s (85%)** |

**Key Insight**: Batch mode provides **10x reduction in Excel sessions** (23 → 2)!

---

## Root Cause Analysis

### Issue 1: LLM Not Using Batch Mode

**Problem**: The LLM performs repetitive operations without batching, causing massive Excel session overhead.

**Evidence from workflow hints**:
```json
{
  "WorkflowHint": "Load-to-data-model configured. For configuring multiple queries, use begin_excel_batch."
}
```

The MCP server **tells the LLM to use batch mode**, but the LLM ignores or doesn't understand this guidance.

**Possible reasons**:
1. **Hint appears AFTER the operation** - Too late to influence next call
2. **No clear example in prompts** - LLM doesn't know the exact pattern
3. **No urgency/penalty** - Hint sounds optional ("For configuring multiple...")
4. **Batch mode not mentioned in tool descriptions** - Only appears in runtime hints

---

### Issue 2: API Design - Parameter Split Responsibility

The `excel_parameter` API has two separate responsibilities:

1. **Named Range Management** (create/update/delete named ranges)
2. **Cell Value Management** (set/get values in cells that named ranges point to)

This dual responsibility creates confusion:
- **`create`** - Creates named range pointing to a cell reference
- **`set`** - Writes value to the cell that the named range points to

### Missing Documentation

The current documentation doesn't show:
- ✅ The relationship between `create` (cell reference) and `set` (cell value)
- ✅ That `range-set-values` can populate cells BEFORE creating named ranges
- ✅ Best practices for bulk parameter creation
- ✅ Performance implications of multiple `set` calls

### No Bulk Operations

Current API doesn't support:
- ❌ Bulk create multiple named ranges in one call
- ❌ Create + set value in single atomic operation
- ❌ Batch parameter definitions from JSON/CSV

---

## Recommendations

### Priority 1: Add loadDestination Parameter (CRITICAL - Best LLM UX)

**Replace** `loadToWorksheet: bool` with `loadDestination: string` for clearer intent.

#### Step 1: Update Core ImportAsync

```csharp
// File: src/ExcelMcp.Core/Commands/PowerQueryCommands.cs
public async Task<OperationResult> ImportAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile,
    string loadDestination = "worksheet",  // ✅ NEW: explicit destination (default: worksheet - what users expect!)
    string? worksheetName = null)
{
    // Import M code first
    var importResult = await ImportMCode(batch, queryName, mCodeFile);
    if (!importResult.Success) return importResult;
    
    // Configure loading based on destination
    return loadDestination.ToLowerInvariant() switch
    {
        "worksheet" => await SetLoadToTableAsync(batch, queryName, worksheetName),
        "data-model" => await SetLoadToDataModelAsync(batch, queryName),
        "both" => await SetLoadToBothAsync(batch, queryName, worksheetName),
        "connection-only" => new OperationResult { Success = true }, // No further action
        _ => throw new ArgumentException($"Invalid loadDestination: {loadDestination}")
    };
}
```

#### Step 2: Update MCP Tool

```csharp
// File: src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs
[Description(@"Load destination for imported query. Options:
  - 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
  - 'data-model': Load to Power Pivot Data Model (for DAX measures/relationships)
  - 'both': Load to both worksheet AND Data Model
  - 'connection-only': Don't load data (M code imported but not executed - advanced use only)
  Default: 'worksheet'")]
string? loadDestination = null,

[Description("Target worksheet name (optional when loadDestination is 'worksheet' or 'both' - auto-generated if omitted)")]
string? targetSheet = null,
```

#### Step 3: Update Prompts

```csharp
// File: src/ExcelMcp.McpServer/Prompts/ExcelPowerQueryDataModelPrompts.cs
// Update examples to show:
excel_powerquery({ 
  action: "import", 
  queryName: "Sales", 
  sourcePath: "sales.pq",
  loadDestination: "data-model"  // ✅ Clear intent!
})
```

**Benefits**:
- ✅ **One call instead of two** - "Import to Data Model" is one operation
- ✅ **Clearer intent** - destination is explicit upfront
- ✅ **Matches mental model** - LLM thinks "import to X" not "import then load to X"
- ✅ **Enum-style parameter** - easier for LLMs than multiple booleans
- ✅ **Backward compatible** - default "connection-only" matches current `loadToWorksheet: false`

**Breaking change considerations**:
- **NEW default behavior**: Imports now load to worksheet by default (matches user expectations)
- Existing `loadToWorksheet: true` → migrate to `loadDestination: "worksheet"` (or omit, it's the new default!)
- Existing `loadToWorksheet: false` → migrate to `loadDestination: "connection-only"` (must be explicit)
- Could support BOTH parameters during transition period

**Effort**: ~6 hours (Core + MCP + update prompts + tests + migration guide)

---

### Priority 2: Provide Helpful Error for Sensitivity Label Issues (CRITICAL - User Education)

**Issue**: Queries using `File.Contents()` fail when source Excel file has Microsoft Purview sensitivity labels (encryption).

**Root Cause**: Power Query **cannot access encrypted Excel files**. Files with sensitivity labels (other than "Public" or "Non-Business") are encrypted and inaccessible to Power Query.

[Microsoft documentation](https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations):
> "Power Query Online is unable to access encrypted Excel files. Since Excel files labeled with sensitivity types other than 'Public' or 'Non-Business' are encrypted, they aren't accessible through Power Query Online."

**Current behavior**:
- Error message is cryptic: "You may not have permissions to apply its required sensitivity label..."
- LLM doesn't understand this is a **Power Query limitation**, not a permissions issue
- LLM tries workarounds that don't solve the root problem

**Recommended Fix**: Detect and provide actionable guidance

```csharp
// In PowerQueryCommands.SetLoadToDataModelAsync (and other load methods)
catch (COMException ex) when (ex.Message.Contains("protected") || ex.Message.Contains("sensitivity label"))
{
    // Parse M code to extract file path
    var mCode = await GetQueryMCodeAsync(batch, queryName);
    var filePath = ExtractFileContentsPath(mCode); // Extract from File.Contents("path")
    
    return new OperationResult
    {
        Success = false,
        ErrorMessage = $@"Source Excel file has Microsoft Purview sensitivity labels (encryption).

Power Query cannot access encrypted Excel files.

SOLUTION: Change sensitivity label to Public
  - Open: {filePath}
  - Click Home tab → Sensitivity button → Select ""Public"" label
  - Save and close
  - Retry: excel_powerquery({{ action: 'set-load-to-data-model', queryName: '{queryName}' }})

Technical details: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations"
    };
}
```

**Helper method**:
```csharp
private static string? ExtractFileContentsPath(string mCode)
{
    // Parse: File.Contents("D:\path\to\file.xlsx")
    var match = Regex.Match(mCode, @"File\.Contents\(""([^""]+)""\)");
    return match.Success ? match.Groups[1].Value : null;
}
```

**Benefits**:
- ✅ **Educates LLM** about Microsoft Purview limitation
- ✅ **Actionable guidance** - specific file path and clear options
- ✅ **Solves root cause** - removes encryption or changes data source
- ✅ **No wasted MCP calls** - LLM knows exactly what to do
- ✅ **Prevents confusion** - explains it's a Power Query limitation, not permissions

**Why NOT automatic fallback**:
- ❌ Automatic load-to-table → create-table → add-to-datamodel **doesn't help**
- ❌ Power Query still can't read the encrypted source file
- ❌ Would hide the real problem from user
- ✅ User **must** either remove label OR change M code pattern

**Effort**: ~2 hours (error detection, M code parsing, helpful error message)

---

**Issue**: Queries using `File.Contents()` to read external Excel files fail when loading to Data Model with:

```
[DataSource.Error] We can't load data from this source because it's protected. 
You may not have permissions to apply its required sensitivity label, or the 
current workbook may contain features that prevent Excel from applying protection.
```

**Root Cause**: Excel privacy protection prevents combining:
- Private source (`Excel.CurrentWorkbook()` - named ranges in current workbook)
- External source (`File.Contents()` - external Excel files)
- When loading to Power Pivot Data Model

**Current Workaround** (LLM had to discover manually):
1. Load query to worksheet table first (`set-load-to-table`)
2. Create Excel Table from range (`excel_table create`)
3. Add table to Data Model (`excel_table add-to-datamodel`)

**Effort**: ~2 hours (error detection, M code parsing, helpful error message)

---

### Priority 3: Bulk Parameter Creation (Medium Priority - Efficiency Gain)

**Current Issue**: Creating 5 parameters requires 10 MCP calls (create + set for each)

```typescript
// Current pattern (10 calls):
excel_parameter({ action: "create", parameterName: "Start_Date", value: "Sheet1!$A$1" })
excel_parameter({ action: "set", parameterName: "Start_Date", value: "2025-07-01" })
excel_parameter({ action: "create", parameterName: "Duration_Months", value: "Sheet1!$B$1" })
excel_parameter({ action: "set", parameterName: "Duration_Months", value: "12" })
// ... 6 more calls
```

**Recommended Fix**: Add `create-bulk` action

```typescript
// Proposed pattern (1 call):
excel_parameter({
  action: "create-bulk",
  parameters: [
    { name: "Start_Date", reference: "Sheet1!$A$1", value: "2025-07-01" },
    { name: "Duration_Months", reference: "Sheet1!$B$1", value: 12 },
    { name: "ProjectRootDirectory", reference: "Sheet1!$C$1", value: "D:\\source\\repos\\cp_toolkit" },
    { name: "Region", reference: "Sheet1!$D$1", value: "West Europe" },
    { name: "vCPUs_Required", reference: "Sheet1!$E$1", value: 8 }
  ]
})
```

**Implementation**:
```csharp
// File: src/ExcelMcp.Core/Commands/ParameterCommands.cs
public async Task<OperationResult> CreateBulkAsync(
    IExcelBatch batch, 
    List<ParameterDefinition> parameters)
{
    foreach (var param in parameters)
    {
        var createResult = await CreateAsync(batch, param.Name, param.Reference);
        if (!createResult.Success) return createResult;
        
        if (param.Value != null)
        {
            var setResult = await SetAsync(batch, param.Name, param.Value);
            if (!setResult.Success) return setResult;
        }
    }
    return new OperationResult { Success = true };
}

public class ParameterDefinition
{
    public string Name { get; set; }
    public string Reference { get; set; }
    public object? Value { get; set; }
}
```

**Benefits**:
- ✅ 10 calls → 1 call (90% reduction)
- ✅ Single Excel session instead of 10
- ✅ Matches LLM mental model ("create these parameters" = one action)
- ✅ Transaction-like: all succeed or all fail

**Effort**: ~3 hours (Core + MCP wrapper + tests)

---

### Summary: Recommended Implementation Order

Given that **backwards compatibility can be broken** (API is for LLMs):

1. **Priority 1: Add loadDestination parameter** (~6 hours)
   - Replace `loadToWorksheet: bool` with `loadDestination: string`
   - Values: "connection-only" (default), "worksheet", "data-model", "both"
   - Clearest API for LLM mental model
   - BREAKING CHANGE: Existing code needs migration

2. **Priority 2: Helpful error for sensitivity labels** (~2 hours)
   - Detect error pattern (contains "protected" or "sensitivity label")
   - Parse M code to extract File.Contents() file path
   - Provide actionable guidance: remove label OR change M code
   - Educates LLM about Power Query limitation (encrypted files)
   - NO BREAKING CHANGES

3. **Priority 3: Bulk parameter creation** (~3 hours)
   - Add create-bulk action
   - 90% reduction in parameter creation calls
   - NO BREAKING CHANGES (additive API)

**Total effort**: ~11 hours to fix all three issues

**Impact**:
- ✅ **50% fewer MCP calls** (import + load in one operation)
- ✅ **Clear guidance on sensitivity labels** (LLM knows to remove label or change M code)
- ✅ **90% fewer parameter calls** (bulk creation)
- ✅ **Clearer mental model** for LLMs (explicit destinations, atomic operations)

---

### Migration Guide (Priority 1 Breaking Change)

**Old Code** (using loadToWorksheet):
```typescript
// Import connection-only
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq", loadToWorksheet: false })

// Import to worksheet
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq", loadToWorksheet: true, targetSheet: "Data" })
```

**New Code** (using loadDestination):
```typescript
// Import to worksheet (NEW DEFAULT - most common case, just omit parameter!)
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq" })
// OR explicitly:
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq", loadDestination: "worksheet", targetSheet: "Data" })

// Import to Data Model (NEW - one call!)
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq", loadDestination: "data-model" })

// Import to both (NEW)
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq", loadDestination: "both", targetSheet: "Data" })

// Import connection-only (advanced - must be explicit)
excel_powerquery({ action: "import", queryName: "Sales", sourcePath: "sales.pq", loadDestination: "connection-only" })
```

**Automatic Migration Script** (for existing .pq workflows):
```typescript
// Search pattern: loadToWorksheet: true
// Replace with: (omit parameter - "worksheet" is now the default!)

// Search pattern: loadToWorksheet: false
// Replace with: loadDestination: "connection-only" (MUST be explicit now)
```

---

If breaking changes are undesirable, add `loadDestination` as NEW parameter alongside `loadToWorksheet`:

```csharp
bool? loadToWorksheet = null,  // ✅ Keep for backward compat
string? loadDestination = null,  // ✅ NEW - preferred way

// Validation logic:
if (loadDestination != null && loadToWorksheet != null)
    throw new McpException("Cannot specify both loadToWorksheet and loadDestination");

// Precedence: loadDestination takes priority if specified
string actualDestination = loadDestination ?? (loadToWorksheet == true ? "worksheet" : loadToWorksheet == false ? "connection-only" : "worksheet");
```

**Benefits**:
- ✅ Zero breaking changes
- ✅ New LLM code uses clearer `loadDestination`
- ✅ Old code continues working with `loadToWorksheet`
- ✅ Gradual migration path

**Drawbacks**:
- ⚠️ Two ways to do the same thing (API complexity)
- ⚠️ LLM might get confused which to use

---

### Priority 3: Document Current Two-Step Pattern Better (Lowest Priority)

If API changes are too risky, at minimum update prompts to be MORE EXPLICIT:

```csharp
public static ChatMessage PowerQueryDataModelGuide()
{
    return new ChatMessage(ChatRole.User, @"
CRITICAL: Import action ONLY imports M code - it does NOT load data!

WRONG (what LLMs naturally try):
excel_powerquery({ action: 'import', loadToWorksheet: false })  ← Only imports, doesn't load to Data Model

RIGHT (two-step pattern required):
excel_powerquery({ action: 'import', loadToWorksheet: false })  ← Step 1: Import M code
excel_powerquery({ action: 'set-load-to-data-model' })          ← Step 2: Load to Data Model

BETTER (use batch mode):
batch = begin_excel_batch({ excelPath })
excel_powerquery({ action: 'import', loadToWorksheet: false, batchId })
excel_powerquery({ action: 'set-load-to-data-model', batchId })
commit_excel_batch({ batchId, save: true })
```

**BUT** this doesn't fix the fundamental UX issue - it just documents the workaround.

---

**Implementation**: Add `loadToDataModel` parameter to import action

#### Step 1: Update Core ImportAsync

```csharp
// File: src/ExcelMcp.Core/Commands/PowerQueryCommands.cs
public async Task<OperationResult> ImportAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile, 
    bool loadToWorksheet = false,      // ✅ Default changed to false
    string? worksheetName = null,
    bool loadToDataModel = false)      // ✅ NEW parameter
{
    // Import M code
    var importResult = await ImportMCode(batch, queryName, mCodeFile);
    
    if (!importResult.Success) return importResult;
    
    // Configure loading based on parameters
    if (loadToWorksheet && loadToDataModel)
    {
        // Load to both worksheet AND Data Model
        return await SetLoadToBothAsync(batch, queryName, worksheetName);
    }
    else if (loadToWorksheet)
    {
        // Load to worksheet only (current behavior)
        return await SetLoadToTableAsync(batch, queryName, worksheetName);
    }
    else if (loadToDataModel)
    {
        // Load to Data Model only (NEW)
        return await SetLoadToDataModelAsync(batch, queryName);
    }
    else
    {
        // Connection-only (current behavior when loadToWorksheet=false)
        return await SetConnectionOnlyAsync(batch, queryName);
    }
}
```

#### Step 2: Update MCP Tool

```csharp
// File: src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs
public static async Task<string> ExcelPowerQuery(
    [Required] string action,
    [Required] string excelPath,
    string? queryName = null,
    string? sourcePath = null,
    
    [Description("Load query data to worksheet for validation (default: false). Cannot be used with loadToDataModel.")]
    bool? loadToWorksheet = null,
    
    [Description("Load query data to Power Pivot Data Model (default: false). Cannot be used with loadToWorksheet unless loadToBoth is used.")]
    bool? loadToDataModel = null,
    
    [Description("Load query data to BOTH worksheet and Data Model (default: false). Overrides loadToWorksheet and loadToDataModel.")]
    bool? loadToBoth = null,
    
    string? batchId = null)
{
    // Validation
    if (loadToBoth == true && (loadToWorksheet == true || loadToDataModel == true))
        throw new McpException("Cannot specify loadToBoth with loadToWorksheet or loadToDataModel");
    
    // Call Core with appropriate parameters
    await commands.ImportAsync(
        batch, 
        queryName, 
        sourcePath, 
        loadToWorksheet: loadToBoth == true || loadToWorksheet == true,
        worksheetName: targetSheet,
        loadToDataModel: loadToBoth == true || loadToDataModel == true);
}
```

#### Step 3: Update Documentation

```markdown
## excel_powerquery import action

Import Power Query M code from .pq file.

**Parameters**:
- `loadToWorksheet` (optional): Load to worksheet table (default: false)
- `loadToDataModel` (optional): Load to Power Pivot Data Model (default: false)  
- `loadToBoth` (optional): Load to both worksheet and Data Model (default: false)

**Examples**:
```typescript
// Import to Data Model (typical Power Pivot workflow)
excel_powerquery({ 
  action: "import", 
  queryName: "Sales", 
  sourcePath: "sales.pq",
  loadToDataModel: true 
})

// Import to worksheet (validation/debugging)
excel_powerquery({ 
  action: "import", 
  queryName: "Sales", 
  sourcePath: "sales.pq",
  loadToWorksheet: true 
})

// Import to both (full setup)
excel_powerquery({ 
  action: "import", 
  queryName: "Sales", 
  sourcePath: "sales.pq",
  loadToBoth: true,
  targetSheet: "SalesData"
})

// Connection-only (default if all omitted)
excel_powerquery({ 
  action: "import", 
  queryName: "Sales", 
  sourcePath: "sales.pq"
})
```

**Benefits**:
- ✅ Cuts Data Model workflow calls in HALF (4 instead of 8 for 4 queries)
- ✅ Matches user mental model ("import to Data Model")
- ✅ Consistent with worksheet loading behavior
- ✅ Clearer intent in code

**Effort**: ~4 hours (Core + MCP + tests + docs)

---

### Priority 2: Make Batch Mode Obvious (HIGH - Documentation)

**Update tool [Description] attributes** to prominently mention batch mode:

```csharp
// BEFORE
[Description("Manage Power Query M code and data loading. Supports: list, view, import, ...")]

// AFTER
[Description(@"Manage Power Query M code and data loading. 

⚡ PERFORMANCE TIP: For multiple operations on same file, use begin_excel_batch first, 
then pass batchId parameter to avoid repeated Excel launches (75-90% faster).

Example - Loading 4 queries to Data Model:
  batch = begin_excel_batch({ excelPath })
  excel_powerquery({ action: 'import', loadToDataModel: true, batchId, ... })  // 4x
  commit_excel_batch({ batchId, save: true })

Supports: list, view, import, export, update, refresh, delete, set-load-to-table, 
set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config.")]
```

**Apply to ALL tools**: ExcelPowerQueryTool, ExcelParameterTool, ExcelRangeTool, ExcelTableTool, etc.

**Files to update**:
- `src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs`
- `src/ExcelMcp.McpServer/Tools/ExcelParameterTool.cs`
- `src/ExcelMcp.McpServer/Tools/ExcelRangeTool.cs`
- `src/ExcelMcp.McpServer/Tools/ExcelTableTool.cs`
- `src/ExcelMcp.McpServer/Tools/ExcelDataModelTool.cs`
- All other tool files

**Benefits**:
- LLM sees batch mode info BEFORE making first call
- Clear performance numbers (75-90% faster)
- Concrete example showing the pattern
- No code changes needed, just documentation

**Effort**: ~2 hours (update all tool descriptions, test MCP protocol response)

---

### Priority 2: Enhance Workflow Hints (Today)

### Priority 2: Enhance Workflow Hints (Today)

**Current hints are too passive**:
```json
{
  "WorkflowHint": "For configuring multiple queries, use begin_excel_batch."
}
```

**Make hints more actionable and urgent**:
```json
{
  "WorkflowHint": "⚠️ PERFORMANCE WARNING: Not using batch mode. If you have more operations on this file, you should use begin_excel_batch to avoid slow repeated Excel launches."
}
```

**Add detection for repeated operations**:
```csharp
// In ExcelToolsBase.WithBatchAsync()
```

---

## Implementation Plan

### Overview

**Total Effort**: ~11 hours (6 + 2 + 3)  
**Priority Order**: Fix API design flaw first, then add helpful errors, then add bulk operations  
**Breaking Changes**: Priority 1 only (acceptable - API designed for LLMs)

---

### Priority 1: Add loadDestination Parameter (~6 hours)

**Goal**: Replace boolean `loadToWorksheet` with enum-style `loadDestination` string parameter.

**Default Behavior Change**: Imports now load to worksheet by default (matches user expectations).

#### Files to Modify

1. **Core Commands** - `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`
2. **MCP Tool** - `src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs`
3. **Prompts** - `src/ExcelMcp.McpServer/Prompts/ExcelPowerQueryDataModelPrompts.cs`
4. **Tests** - `tests/ExcelMcp.Core.Tests/PowerQueryCommandsTests.cs`
5. **MCP Tests** - `tests/ExcelMcp.McpServer.Tests/ExcelPowerQueryToolTests.cs`
6. **Documentation** - `docs/COMMANDS.md`

#### Step 1.1: Update Core ImportAsync (~2 hours)

**File**: `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`

**Current Signature** (line ~551):
```csharp
public async Task<OperationResult> ImportAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile, 
    bool loadToWorksheet = true,
    string? worksheetName = null)
```

**New Signature**:
```csharp
public async Task<OperationResult> ImportAsync(
    IExcelBatch batch, 
    string queryName, 
    string mCodeFile, 
    string loadDestination = "worksheet",  // NEW: explicit destination
    string? worksheetName = null)
{
    // Import M code first
    var importResult = await ImportMCodeAsync(batch, queryName, mCodeFile);
    if (!importResult.Success) return importResult;
    
    // Configure loading based on destination
    return loadDestination.ToLowerInvariant() switch
    {
        "worksheet" => await SetLoadToTableAsync(batch, queryName, worksheetName),
        "data-model" => await SetLoadToDataModelAsync(batch, queryName),
        "both" => await SetLoadToBothAsync(batch, queryName, worksheetName),
        "connection-only" => new OperationResult { Success = true },
        _ => throw new ArgumentException(
            $"Invalid loadDestination: '{loadDestination}'. " +
            $"Valid values: 'worksheet', 'data-model', 'both', 'connection-only'",
            nameof(loadDestination))
    };
}
```

**Changes Required**:
- ✅ Replace `bool loadToWorksheet` parameter with `string loadDestination`
- ✅ Add validation for valid destination values
- ✅ Implement switch expression for routing
- ✅ Ensure `SetLoadToBothAsync` exists (may need to implement)

#### Step 1.2: Update MCP Tool (~1.5 hours)

**File**: `src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs`

**Current Parameter** (around line 150-155):
```csharp
[Description("Automatically load query data to worksheet for validation (default: true). When false, creates connection-only query without validation.")]
bool? loadToWorksheet = null,
```

**New Parameter**:
```csharp
[Description(@"Load destination for imported query. Options:
  - 'worksheet': Load to worksheet as table (DEFAULT - users can see/validate data)
  - 'data-model': Load to Power Pivot Data Model (for DAX measures/relationships)
  - 'both': Load to both worksheet AND Data Model
  - 'connection-only': Don't load data (M code imported but not executed - advanced use only)
Default: 'worksheet'")]
string? loadDestination = null,

[Description("Target worksheet name (optional when loadDestination is 'worksheet' or 'both' - auto-generated if omitted)")]
string? targetSheet = null,
```

**Update Import Action Handler**:
```csharp
case "import":
    ValidateParameters(queryName, sourcePath);
    
    // Use default if not specified
    var destination = loadDestination ?? "worksheet";
    
    var importResult = await _commands.ImportAsync(
        batch, 
        queryName!, 
        sourcePath!,
        destination,
        targetSheet);
    
    return SerializeResult(importResult);
```

**Backward Compatibility Option** (if desired):
```csharp
// Support BOTH old and new parameters during transition
bool? loadToWorksheet = null,  // Deprecated but supported
string? loadDestination = null,

// In handler:
if (loadToWorksheet != null && loadDestination != null)
    throw new McpException("Cannot specify both loadToWorksheet and loadDestination");

var destination = loadDestination ?? 
    (loadToWorksheet == true ? "worksheet" : 
     loadToWorksheet == false ? "connection-only" : 
     "worksheet");
```

#### Step 1.3: Update Prompts (~30 minutes)

**File**: `src/ExcelMcp.McpServer/Prompts/ExcelPowerQueryDataModelPrompts.cs`

**Update Examples**:
```csharp
public static ChatMessage PowerQueryDataModelGuide()
{
    return new ChatMessage(ChatRole.User, @"
# Power Query Data Model Loading - Best Practices

## Import with Load Destination (Recommended)

Import Power Query and specify where to load data in ONE call:

```typescript
// Load to Data Model (Power Pivot workflows)
excel_powerquery({ 
  action: 'import', 
  queryName: 'Sales', 
  sourcePath: 'sales.pq',
  loadDestination: 'data-model'  // ✅ One operation!
})

// Load to worksheet (default - most common)
excel_powerquery({ 
  action: 'import', 
  queryName: 'Sales', 
  sourcePath: 'sales.pq'
  // loadDestination: 'worksheet' is default - can omit!
})

// Load to both worksheet AND Data Model
excel_powerquery({ 
  action: 'import', 
  queryName: 'Sales', 
  sourcePath: 'sales.pq',
  loadDestination: 'both',
  targetSheet: 'SalesData'
})

// Connection-only (advanced - no data loaded)
excel_powerquery({ 
  action: 'import', 
  queryName: 'Helper', 
  sourcePath: 'helper.pq',
  loadDestination: 'connection-only'
})
```

## Batch Mode for Multiple Queries

When importing multiple queries, use batch mode:

```typescript
batch = begin_excel_batch({ excelPath: 'model.xlsx' })

// Import 4 queries to Data Model in one session
excel_powerquery({ action: 'import', queryName: 'Q1', sourcePath: 'q1.pq', loadDestination: 'data-model', batchId })
excel_powerquery({ action: 'import', queryName: 'Q2', sourcePath: 'q2.pq', loadDestination: 'data-model', batchId })
excel_powerquery({ action: 'import', queryName: 'Q3', sourcePath: 'q3.pq', loadDestination: 'data-model', batchId })
excel_powerquery({ action: 'import', queryName: 'Q4', sourcePath: 'q4.pq', loadDestination: 'data-model', batchId })

commit_excel_batch({ batchId, save: true })
// Result: 1 Excel session instead of 4 (75% faster!)
```
");
}
```

#### Step 1.4: Update Tests (~1.5 hours)

**File**: `tests/ExcelMcp.Core.Tests/PowerQueryCommandsTests.cs`

**Add New Tests**:
```csharp
[Fact]
public async Task ImportAsync_WithWorksheetDestination_LoadsToWorksheet()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    var result = await _commands.ImportAsync(
        batch, 
        "TestQuery", 
        _testQueryFile,
        loadDestination: "worksheet");
    
    Assert.True(result.Success);
    // Verify query loaded to worksheet
}

[Fact]
public async Task ImportAsync_WithDataModelDestination_LoadsToDataModel()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    var result = await _commands.ImportAsync(
        batch, 
        "TestQuery", 
        _testQueryFile,
        loadDestination: "data-model");
    
    Assert.True(result.Success);
    // Verify query loaded to Data Model
}

[Fact]
public async Task ImportAsync_WithBothDestination_LoadsToBoth()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    var result = await _commands.ImportAsync(
        batch, 
        "TestQuery", 
        _testQueryFile,
        loadDestination: "both",
        worksheetName: "Data");
    
    Assert.True(result.Success);
    // Verify query loaded to both worksheet AND Data Model
}

[Fact]
public async Task ImportAsync_WithConnectionOnlyDestination_NoDataLoaded()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    var result = await _commands.ImportAsync(
        batch, 
        "TestQuery", 
        _testQueryFile,
        loadDestination: "connection-only");
    
    Assert.True(result.Success);
    // Verify query imported but not loaded
}

[Fact]
public async Task ImportAsync_DefaultDestination_LoadsToWorksheet()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    // Omit loadDestination - should default to "worksheet"
    var result = await _commands.ImportAsync(
        batch, 
        "TestQuery", 
        _testQueryFile);
    
    Assert.True(result.Success);
    // Verify query loaded to worksheet (default)
}

[Fact]
public async Task ImportAsync_InvalidDestination_ThrowsArgumentException()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    await Assert.ThrowsAsync<ArgumentException>(async () =>
        await _commands.ImportAsync(
            batch, 
            "TestQuery", 
            _testQueryFile,
            loadDestination: "invalid"));
}
```

**File**: `tests/ExcelMcp.McpServer.Tests/ExcelPowerQueryToolTests.cs`

**Add MCP Protocol Tests**:
```csharp
[Fact]
public async Task Import_WithDataModelDestination_CallsCoreCorrectly()
{
    var result = await ExcelPowerQueryTool.ExcelPowerQuery(
        action: "import",
        excelPath: _testFile,
        queryName: "TestQuery",
        sourcePath: _testQueryFile,
        loadDestination: "data-model");
    
    // Verify result contains success
    Assert.Contains("\"Success\":true", result);
}
```

#### Step 1.5: Update Documentation (~30 minutes)

**File**: `docs/COMMANDS.md`

**Update Power Query Section**:
```markdown
**pq-import** - Create or import query from file

```powershell
excelcli pq-import <file.xlsx> <query-name> <source.pq> [--destination <worksheet|data-model|both|connection-only>] [--target-sheet <sheet-name>]
```

Import a Power Query from an M code file with optional load destination.

**Load Destination Options**:
- `worksheet` (DEFAULT): Load to worksheet as table (users can see/validate data)
- `data-model`: Load to Power Pivot Data Model (for DAX measures/relationships)
- `both`: Load to both worksheet AND Data Model
- `connection-only`: Don't load data (M code imported but not executed - advanced use only)

**Examples**:
```powershell
# Import to worksheet (default behavior)
excelcli pq-import data.xlsx "Sales" sales.pq

# Import to Data Model
excelcli pq-import data.xlsx "Sales" sales.pq --destination data-model

# Import to both
excelcli pq-import data.xlsx "Sales" sales.pq --destination both --target-sheet SalesData

# Connection-only (advanced)
excelcli pq-import data.xlsx "Helper" helper.pq --destination connection-only
```

**BREAKING CHANGE from v1.x**:
- Old: `--load-to-worksheet false` → New: `--destination connection-only`
- Old: `--load-to-worksheet true` (or omit) → New: (omit - worksheet is default)
```

#### Step 1.6: Migration Guide (~30 minutes)

**Create**: `docs/MIGRATION-GUIDE-V2.md`

```markdown
# Migration Guide: v1.x → v2.0

## Breaking Change: loadDestination Parameter

### Overview

The `loadToWorksheet: bool` parameter has been replaced with `loadDestination: string` for clearer intent and better LLM UX.

### Migration

**Before (v1.x)**:
```typescript
// Load to worksheet
excel_powerquery({ action: "import", loadToWorksheet: true })

// Connection-only
excel_powerquery({ action: "import", loadToWorksheet: false })
```

**After (v2.0)**:
```typescript
// Load to worksheet (DEFAULT - can omit!)
excel_powerquery({ action: "import" })
// Or explicitly:
excel_powerquery({ action: "import", loadDestination: "worksheet" })

// Connection-only (must be explicit now)
excel_powerquery({ action: "import", loadDestination: "connection-only" })

// NEW: Load to Data Model (one call!)
excel_powerquery({ action: "import", loadDestination: "data-model" })

// NEW: Load to both
excel_powerquery({ action: "import", loadDestination: "both", targetSheet: "Data" })
```

### Automatic Migration Script

```powershell
# Find all old patterns
Get-ChildItem -Recurse -Include *.ts,*.js,*.pq | 
    Select-String "loadToWorksheet: true" |
    ForEach-Object { $_.Path }

# Replace manually:
# loadToWorksheet: true → (remove parameter - worksheet is now default)
# loadToWorksheet: false → loadDestination: "connection-only"
```

### Benefits

- ✅ **Clearer intent**: Explicitly states where data goes
- ✅ **One operation**: Import to Data Model in single call
- ✅ **Better defaults**: Worksheet loading matches user expectations
- ✅ **Enum-style**: Easier for LLMs than boolean combinations
```

---

### Priority 2: Helpful Error for Sensitivity Labels (~2 hours)

**Goal**: Detect Microsoft Purview sensitivity label errors and provide actionable guidance.

#### Files to Modify

1. **Core Commands** - `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`
2. **Tests** - `tests/ExcelMcp.Core.Tests/PowerQueryCommandsTests.cs`

#### Step 2.1: Implement Error Detection (~1 hour)

**File**: `src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`

**Add Helper Method**:
```csharp
/// <summary>
/// Extracts file path from File.Contents() in M code
/// </summary>
private static string? ExtractFileContentsPath(string mCode)
{
    // Parse: File.Contents("D:\path\to\file.xlsx")
    var match = Regex.Match(mCode, @"File\.Contents\s*\(\s*""([^""]+)""\s*\)");
    return match.Success ? match.Groups[1].Value : null;
}
```

**Update SetLoadToDataModelAsync** (around line 400-500):
```csharp
public async Task<OperationResult> SetLoadToDataModelAsync(
    IExcelBatch batch, 
    string queryName)
{
    return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
    {
        try
        {
            // Existing implementation...
            dynamic connection = FindConnection(ctx.Book, queryName);
            // ... load to Data Model logic ...
            
            return new OperationResult { Success = true };
        }
        catch (COMException ex) when (
            ex.Message.Contains("protected", StringComparison.OrdinalIgnoreCase) || 
            ex.Message.Contains("sensitivity label", StringComparison.OrdinalIgnoreCase))
        {
            // Get M code to extract file path
            var mCode = await GetQueryMCodeAsync(batch, queryName);
            var filePath = ExtractFileContentsPath(mCode);
            
            var filePathInfo = !string.IsNullOrEmpty(filePath) 
                ? $"\n  File: {filePath}\n" 
                : "\n";
            
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $@"Source Excel file has Microsoft Purview sensitivity labels (encryption).

Power Query cannot access encrypted Excel files. Choose one option:

OPTION 1 (Recommended): Set sensitivity label to Public
  - Open:{filePathInfo}  - File → Info → Sensitivity → Change label to ""Public""
  - Save and close
  - Retry: excel_powerquery({{ action: 'set-load-to-data-model', queryName: '{queryName}' }})

OPTION 2: Modify M code to use different data source
  - Replace File.Contents() with Excel.CurrentWorkbook() if data is in same workbook
  - Export source data to CSV and use Csv.Document()
  - Use ODBC or SQL connection if source is a database
  - Use: excel_powerquery({{ action: 'update', queryName: '{queryName}', sourcePath: 'modified.pq' }})

Technical details: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations"
            };
        }
    });
}
```

**Also Update** (same error handling):
- `SetLoadToTableAsync`
- `SetLoadToBothAsync`
- Any other methods that trigger query refresh

#### Step 2.2: Add Tests (~1 hour)

**File**: `tests/ExcelMcp.Core.Tests/PowerQueryCommandsTests.cs`

**Add Test**:
```csharp
[Fact]
public async Task SetLoadToDataModelAsync_SensitivityLabelError_ReturnsHelpfulMessage()
{
    // Setup: Create query with File.Contents pointing to labeled file
    // (This test may need manual setup or mocking)
    
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    // Create query that will fail due to sensitivity label
    var mCode = @"
let
    Source = File.Contents(""D:\labeled\file.xlsx""),
    Excel = Excel.Workbook(Source)
in
    Excel";
    
    File.WriteAllText(_testQueryFile, mCode);
    await _commands.ImportAsync(batch, "LabeledQuery", _testQueryFile, "connection-only");
    
    // Attempt to load - should fail with helpful message
    var result = await _commands.SetLoadToDataModelAsync(batch, "LabeledQuery");
    
    Assert.False(result.Success);
    Assert.Contains("Microsoft Purview sensitivity labels", result.ErrorMessage);
    Assert.Contains("OPTION 1", result.ErrorMessage);
    Assert.Contains("OPTION 2", result.ErrorMessage);
    Assert.Contains("D:\\labeled\\file.xlsx", result.ErrorMessage);
}

[Fact]
public void ExtractFileContentsPath_ValidMCode_ReturnsPath()
{
    var mCode = @"File.Contents(""D:\data\source.xlsx"")";
    var path = ExtractFileContentsPath(mCode);
    
    Assert.Equal(@"D:\data\source.xlsx", path);
}

[Fact]
public void ExtractFileContentsPath_NoFileContents_ReturnsNull()
{
    var mCode = @"Excel.CurrentWorkbook()";
    var path = ExtractFileContentsPath(mCode);
    
    Assert.Null(path);
}
```

---

### Priority 3: Bulk Parameter Creation (~3 hours)

**Goal**: Add `create-bulk` action to create multiple named ranges with values in one call.

#### Files to Modify

1. **Core Models** - `src/ExcelMcp.Core/Models/ParameterDefinition.cs` (new file)
2. **Core Commands** - `src/ExcelMcp.Core/Commands/ParameterCommands.cs`
3. **MCP Tool** - `src/ExcelMcp.McpServer/Tools/ExcelParameterTool.cs`
4. **Tests** - `tests/ExcelMcp.Core.Tests/ParameterCommandsTests.cs`
5. **MCP Tests** - `tests/ExcelMcp.McpServer.Tests/ExcelParameterToolTests.cs`
6. **Documentation** - `docs/COMMANDS.md`

#### Step 3.1: Create Model Class (~15 minutes)

**File**: `src/ExcelMcp.Core/Models/ParameterDefinition.cs` (NEW)

```csharp
namespace ExcelMcp.Core.Models;

/// <summary>
/// Defines a named range parameter with optional initial value
/// </summary>
public class ParameterDefinition
{
    /// <summary>
    /// Name of the named range
    /// </summary>
    public required string Name { get; set; }
    
    /// <summary>
    /// Cell reference (e.g., "Sheet1!$A$1")
    /// </summary>
    public required string Reference { get; set; }
    
    /// <summary>
    /// Optional initial value to set in the cell
    /// </summary>
    public object? Value { get; set; }
}
```

#### Step 3.2: Update Core Commands (~1 hour)

**File**: `src/ExcelMcp.Core/Commands/ParameterCommands.cs`

**Add Method**:
```csharp
/// <summary>
/// Creates multiple named ranges with optional initial values
/// </summary>
public async Task<OperationResult> CreateBulkAsync(
    IExcelBatch batch, 
    List<ParameterDefinition> parameters)
{
    if (parameters == null || parameters.Count == 0)
        return new OperationResult 
        { 
            Success = false, 
            ErrorMessage = "No parameters provided" 
        };
    
    foreach (var param in parameters)
    {
        // Validate
        if (string.IsNullOrWhiteSpace(param.Name))
            return new OperationResult 
            { 
                Success = false, 
                ErrorMessage = $"Parameter name cannot be empty" 
            };
        
        if (string.IsNullOrWhiteSpace(param.Reference))
            return new OperationResult 
            { 
                Success = false, 
                ErrorMessage = $"Parameter '{param.Name}' reference cannot be empty" 
            };
        
        // Create named range
        var createResult = await CreateAsync(batch, param.Name, param.Reference);
        if (!createResult.Success) 
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create parameter '{param.Name}': {createResult.ErrorMessage}"
            };
        
        // Set value if provided
        if (param.Value != null)
        {
            var setResult = await SetAsync(batch, param.Name, param.Value);
            if (!setResult.Success) 
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set value for parameter '{param.Name}': {setResult.ErrorMessage}"
                };
        }
    }
    
    return new OperationResult 
    { 
        Success = true,
        Message = $"Created {parameters.Count} parameter(s)"
    };
}
```

#### Step 3.3: Update MCP Tool (~1 hour)

**File**: `src/ExcelMcp.McpServer/Tools/ExcelParameterTool.cs`

**Add Parameter**:
```csharp
[Description("Bulk parameter definitions (for create-bulk action only). Array of {name, reference, value?} objects.")]
string? parametersJson = null,
```

**Add Action Handler**:
```csharp
case "create-bulk":
    if (string.IsNullOrWhiteSpace(parametersJson))
        throw new McpException("parametersJson required for create-bulk action");
    
    var parameters = JsonSerializer.Deserialize<List<ParameterDefinition>>(
        parametersJson, 
        new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
    
    if (parameters == null)
        throw new McpException("Invalid parametersJson format");
    
    var bulkResult = await _commands.CreateBulkAsync(batch, parameters);
    return SerializeResult(bulkResult);
```

#### Step 3.4: Add Tests (~45 minutes)

**File**: `tests/ExcelMcp.Core.Tests/ParameterCommandsTests.cs`

```csharp
[Fact]
public async Task CreateBulkAsync_MultipleParameters_CreatesAll()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    var parameters = new List<ParameterDefinition>
    {
        new() { Name = "Param1", Reference = "Sheet1!$A$1", Value = "Value1" },
        new() { Name = "Param2", Reference = "Sheet1!$B$1", Value = 123 },
        new() { Name = "Param3", Reference = "Sheet1!$C$1", Value = DateTime.Now }
    };
    
    var result = await _commands.CreateBulkAsync(batch, parameters);
    
    Assert.True(result.Success);
    
    // Verify all created
    var param1 = await _commands.GetAsync(batch, "Param1");
    Assert.Equal("Value1", param1.Value);
    
    var param2 = await _commands.GetAsync(batch, "Param2");
    Assert.Equal(123, Convert.ToInt32(param2.Value));
}

[Fact]
public async Task CreateBulkAsync_EmptyList_ReturnsError()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    
    var result = await _commands.CreateBulkAsync(batch, new List<ParameterDefinition>());
    
    Assert.False(result.Success);
    Assert.Contains("No parameters provided", result.ErrorMessage);
}
```

**File**: `tests/ExcelMcp.McpServer.Tests/ExcelParameterToolTests.cs`

```csharp
[Fact]
public async Task CreateBulk_ValidParameters_CreatesAll()
{
    var parametersJson = JsonSerializer.Serialize(new[]
    {
        new { name = "Start_Date", reference = "Sheet1!$A$1", value = "2025-07-01" },
        new { name = "Duration", reference = "Sheet1!$B$1", value = 12 }
    });
    
    var result = await ExcelParameterTool.ExcelParameter(
        action: "create-bulk",
        excelPath: _testFile,
        parametersJson: parametersJson);
    
    Assert.Contains("\"Success\":true", result);
}
```

#### Step 3.5: Update Documentation (~15 minutes)

**File**: `docs/COMMANDS.md`

**Add Command**:
```markdown
**param-create-bulk** - Create multiple named ranges with values ✨ **NEW**

```powershell
excelcli param-create-bulk <file.xlsx> <params.json>
```

Create multiple named ranges with optional initial values in a single operation.

**JSON Format**:
```json
[
  { "name": "Start_Date", "reference": "Sheet1!$A$1", "value": "2025-07-01" },
  { "name": "Duration_Months", "reference": "Sheet1!$B$1", "value": 12 },
  { "name": "ProjectRoot", "reference": "Sheet1!$C$1", "value": "D:\\source\\repos" }
]
```

**Example**:
```powershell
# Create params.json with parameter definitions
@"
[
  { "name": "Start_Date", "reference": "Sheet1!$A$1", "value": "2025-07-01" },
  { "name": "Duration_Months", "reference": "Sheet1!$B$1", "value": 36 },
  { "name": "Plan_Name", "reference": "Sheet1!$C$1", "value": "FY26 Consumption Plan" }
]
"@ | Out-File params.json

# Create all parameters at once
excelcli param-create-bulk data.xlsx params.json
```

**Benefits**:
- ✅ 90% reduction in MCP calls (10 calls → 1 call for 5 parameters)
- ✅ Single Excel session instead of 10
- ✅ Atomic operation: all succeed or all fail
```

---

### Testing Strategy

#### Unit Tests
- ✅ All new methods have happy path tests
- ✅ Error cases covered (invalid destinations, empty bulk lists, etc.)
- ✅ Edge cases (null values, special characters in paths)

#### Integration Tests
- ✅ End-to-end workflows with real Excel files
- ✅ Sensitivity label error handling (may require manual file setup)
- ✅ Batch mode + new parameters working together

#### MCP Protocol Tests
- ✅ JSON serialization/deserialization working correctly
- ✅ Tool descriptions accurate
- ✅ Error messages formatted properly for LLM consumption

---

### Rollout Plan

#### Phase 1: Priority 1 (Week 1)
- Day 1-2: Core implementation + tests
- Day 3: MCP tool updates
- Day 4: Documentation + migration guide
- Day 5: Testing + bug fixes

#### Phase 2: Priority 2 (Week 1)
- Day 1: Sensitivity label error detection
- Day 2: Testing + refinement

#### Phase 3: Priority 3 (Week 2)
- Day 1-2: Bulk parameter creation
- Day 3: Testing + documentation

---

### Success Metrics

**After Implementation**:
- ✅ Data Model workflow: 4 calls instead of 8 (50% reduction)
- ✅ Parameter creation: 1 call instead of 10 (90% reduction)
- ✅ Sensitivity label errors: Clear guidance instead of cryptic message
- ✅ Combined workflow: 11 calls instead of 24 (54% reduction)
- ✅ Excel sessions: 3 instead of 23 (87% reduction with batch mode)

---

### Risk Mitigation

**Breaking Changes**:
- ✅ Migration guide provided
- ✅ Clear error messages for old parameter usage
- ✅ Optional backward compatibility mode available

**Sensitivity Label Detection**:
- ⚠️ May require testing with actual labeled files
- ⚠️ Error message patterns may vary by Excel version
- ✅ Graceful degradation if file path can't be extracted

**Bulk Operations**:
- ⚠️ Large parameter lists may hit Excel limits
- ✅ Transaction-like behavior (all-or-nothing)
- ✅ Clear error messages for which parameter failed
if (batchId == null && _recentFiles.Contains(filePath, within: TimeSpan.FromSeconds(30)))
{
    result.PerformanceWarning = "⚠️ Multiple operations detected on same file without batch mode. Consider using begin_excel_batch for 75-90% faster execution.";
}
```

---

### Priority 3: Add Bulk Parameter Creation (This Week)

**New API**: Create multiple parameters atomically

```typescript
excel_parameter({ 
  action: "create-bulk",
  excelPath: "file.xlsx",
  sheetName: "Sheet1",
  parameters: [
    { name: "Start_Date", cell: "A1", value: "2025-07-01" },
    { name: "Duration_Months", cell: "B1", value: "36" },
    { name: "Plan_Name", cell: "C1", value: "FY26 Consumption Plan" },
    { name: "Customer_Name", cell: "D1", value: "PhysicsX" },
    { name: "ProjectRoot", cell: "E1", value: "D:\\source\\repos\\cp_toolkit" }
  ]
})

// Single call replaces 10 calls (or 6 with optimization)
```

**Implementation effort**: ~3 hours (Core + MCP + tests)

---

### Priority 4: Auto-Batch Detection (Future)

**Intelligent batch mode recommendation**:

```typescript
// LLM makes first call without batch
excel_powerquery({ action: "set-load-to-data-model", queryName: "Q1" })

// Server response includes smart suggestion
{
  "success": true,
  "autoGeneratedBatchId": "batch_abc123",
  "message": "💡 Auto-batch mode available: I detected you might be loading multiple queries. Reply with 'use-batch:batch_abc123' and I'll keep this Excel session open for your next operations. This will be 75% faster than separate calls."
}

// LLM can continue with batch
excel_powerquery({ action: "set-load-to-data-model", queryName: "Q2", batchId: "batch_abc123" })
```

**Benefits**:
- ✅ Zero LLM learning curve
- ✅ Server guides LLM to efficient patterns
- ✅ Backward compatible (works without LLM changes)

**Complexity**: Medium-high (requires session state management)

---

**Rationale**: `excel_parameter.set` is redundant with `excel_range.set-values`

**Migration Path**:
```typescript
// OLD (redundant)
excel_parameter({ action: "set", parameterName: "Start_Date", value: "2025-07-01" })

// NEW (use range API)
excel_range({ action: "set-values", sheetName: "Sheet1", rangeAddress: "A1", values: [["2025-07-01"]] })
```

**Benefits**:
- ✅ Clearer separation of concerns
- ✅ Parameter tool = named range lifecycle only
- ✅ Range tool = cell value operations
- ✅ Removes API confusion

**Drawbacks**:
- ❌ Breaking change (requires migration)
- ❌ More verbose for single parameter updates
- ❌ May confuse users who expect parameter API to be self-contained

---

## The File.Contents() Privacy Bug

**Separate Issue**: Inconsistent behavior when loading Power Query to Data Model

| Query | Uses File.Contents()? | Result |
|-------|----------------------|--------|
| Regions | ✅ Yes (master-data\\Azure_Regions.xlsx) | ✅ SUCCESS (unexpected) |
| Milestones | ✅ Yes (plans\\PhysicsX C-Plan.xlsx) | ❌ FAILED (expected) |

**Error Message** (Milestones):
```
[DataSource.Error] We can't load data from this source because it's protected. 
You may not have permissions to apply its required sensitivity label, or the 
current workbook may contain features that prevent Excel from applying protection.
```

**Possible Causes**:
1. **File location difference** - `master-data\` vs `plans\` directory
2. **File size** - 24 rows vs 61 rows
3. **Timing** - Regions loaded before privacy barrier established?
4. **File metadata** - Different Excel features or properties

**Workaround** (always works):
```typescript
// Instead of direct Data Model load (sometimes fails)
excel_powerquery({ action: "set-load-to-data-model", queryName: "Milestones" })

// Use this pattern (always works)
excel_powerquery({ action: "set-load-to-table", queryName: "Milestones", targetSheet: "Milestones" })
excel_table({ action: "create", tableName: "Milestones", sheetName: "Milestones", range: "A1:S62" })
excel_table({ action: "add-to-datamodel", tableName: "Milestones" })
```

**Requires Investigation**:
- Why does Regions succeed?
- Is there a privacy level API we're missing?
- Can we auto-detect File.Contents() and apply workaround?

---

## Immediate Actions

### TODAY (CRITICAL - API Design Fix)

- [ ] **Add loadToDataModel parameter to import** (4 hours) - **HIGHEST PRIORITY**
  - Update `ImportAsync` signature in PowerQueryCommands.cs
  - Add conditional logic for loadToDataModel
  - Update MCP tool parameter
  - Add validation (can't use both worksheet and DataModel without loadToBoth)
  - Write tests for all combinations
  - Update documentation

**Impact**: Cuts LLM calls by 50% for Data Model workflows (8 calls → 4 calls)

---

### THIS WEEK (High Priority - User Experience)

- [ ] **Update ALL MCP tool [Description] attributes** to mention batch mode (2 hours)
  - Add "⚡ PERFORMANCE TIP" section to each tool description
  - Include batch mode example with before/after performance
  - Update: PowerQueryTool, ParameterTool, RangeTool, TableTool, DataModelTool, etc.
  - Test MCP protocol shows updated descriptions

- [ ] **Implement bulk parameter creation** (3 hours)
  - Add `create-bulk` action to excel_parameter
  - Single call to create multiple parameters + values
  - Atomic operation

- [ ] **Investigate File.Contents() privacy bug** (2 hours)
  - Why does Regions succeed but Milestones fail?
  - Research Excel COM privacy APIs
  - Document workaround or fix

---

## Performance Impact

**Real-world scenario** (5 parameters):
- Current: 10 MCP calls, 10 Excel sessions
- Efficient: 6 MCP calls, 6 Excel sessions
- Bulk API: 1 MCP call, 1 Excel session

**Time savings** (estimated):
- Current: ~10 seconds (10 × 1 sec per call)
- Efficient: ~6 seconds (6 × 1 sec per call)
- Bulk API: ~1 second (1 call)

**For larger workbooks** (20 parameters):
- Current: 40 calls, ~40 seconds
- Efficient: 21 calls, ~21 seconds (47% reduction)
- Bulk API: 1 call, ~1 second (97% reduction)

---

## Conclusion

The LLM is using the API correctly but inefficiently due to:
1. ❌ API design that splits parameter creation into 2 steps
2. ❌ Missing documentation on efficient patterns
3. ❌ No bulk operations support

**Recommended immediate action**: Update documentation with efficient pattern  
**Recommended short-term action**: Add `create-bulk` API  
**Recommended long-term action**: Consider deprecating `set` action in favor of range API

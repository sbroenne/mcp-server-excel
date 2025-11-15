# Session API Redesign Specification

**Version:** 1.0
**Status:** Draft
**Date:** 2025-01-13
**Author:** Development Team

## Executive Summary

This specification proposes a **breaking redesign** of ExcelMcp's session API to use intuitive **Open/Save/Close** semantics exclusively. The goal is to eliminate the "batch" concept entirely and remove all cognitive load from LLMs - every operation works through sessions, always. No backwards compatibility, no dual patterns, no decisions about when to batch.

## ⚠️ CRITICAL: Excel COM Threading & Concurrency Limitations

**Excel COM API is fundamentally single-threaded and does NOT support parallel operations.**

### Operations Within a Session are ALWAYS SERIAL

- Each session (`IExcelBatch`/`IExcelSession`) runs on **ONE dedicated STA thread** with **ONE Excel instance**
- Operations are **queued** and executed **sequentially** via `Channel<Func<Task>>`
- Multiple `session.Execute()` calls are processed **one at a time** (never in parallel)
- This is a **COM interop requirement**, not an implementation choice

**Example - Operations are SERIAL, not parallel:**

```csharp
// ❌ These do NOT run in parallel - they are queued serially!
var task1 = session.Execute(ctx => ctx.Book.Worksheets.Add("Sheet1"));  // Queued
var task2 = session.Execute(ctx => ctx.Book.Worksheets.Add("Sheet2"));  // Queued AFTER task1
await Task.WhenAll(task1, task2);  // Still serial execution on STA thread!
```

**Why:** Excel COM requires Single-Threaded Apartment (STA) model. No concurrent access to Excel objects is possible.

### Multiple Sessions = Multiple Excel Processes (Resource Heavy)

**You CAN create multiple sessions for DIFFERENT files:**
- Each session = one `Excel.Application` process (~50-100MB+ memory each)
- Sessions run in **separate processes** (true parallelism between files)
- But **operations within each session remain serial**

**Example - Multiple files (parallel processes, serial operations per file):**

```csharp
// ✅ CORRECT: Multiple files = multiple Excel processes (true parallelism)
var session1 = await manager.CreateSessionAsync("fileA.xlsx");  // Excel process 1
var session2 = await manager.CreateSessionAsync("fileB.xlsx");  // Excel process 2

// These run in parallel (different Excel processes)
var task1 = GetSession(session1).Execute(...);  // Runs in Excel process 1
var task2 = GetSession(session2).Execute(...);  // Runs in Excel process 2
await Task.WhenAll(task1, task2);  // ✅ True parallelism (different processes)

// But within each session, operations are still serial
var task3 = GetSession(session1).Execute(...);  // Queued after task1
var task4 = GetSession(session1).Execute(...);  // Queued after task3
```

**Resource Limits:**
- Each Excel process consumes ~50-100MB+ memory
- Windows desktop machines have finite resources
- **Recommendation:** Limit to 3-5 concurrent sessions for typical desktops

### File Creation Must Be Sequential

**File creation is automatically serialized by the implementation:**

```csharp
// ✅ This is now SAFE - internal lock serializes calls automatically
var tasks = Enumerable.Range(1, 10).Select(i =>
    ExcelSession.CreateNew($"file{i}.xlsx", false, ...));
await Task.WhenAll(tasks);  // Executes sequentially despite Task.WhenAll!

// ✅ This is also safe and more explicit
for (int i = 1; i <= 10; i++)
{
    await ExcelSession.CreateNew($"file{i}.xlsx", false, ...);
}
```

**How it works:** `ExcelSession` uses a static `SemaphoreSlim(1, 1)` to serialize all `CreateNew()` and `CreateNewAsync()` calls. Even if called in parallel, they queue and execute one at a time.

**Why enforced:** Each `CreateNew()` temporarily creates an Excel instance, saves the file, then closes it. Without serialization, parallel creation would spawn many Excel processes simultaneously, causing memory exhaustion.

### SessionManager Prevents Same-File Conflicts

```csharp
// SessionManager enforces one session per file
await manager.CreateSessionAsync("sales.xlsx");  // ✅ OK
await manager.CreateSessionAsync("sales.xlsx");  // ❌ Throws: "File already open in another session"
```

This matches Excel UI behavior (cannot open same file twice).

### Key Takeaways for Implementation

1. **Within-session parallelism is IMPOSSIBLE** - all operations are queued serially on STA thread
2. **Between-sessions parallelism is POSSIBLE** - different files = different processes
3. **File creation is AUTOMATICALLY SERIALIZED** - enforced by SemaphoreSlim lock (prevents resource exhaustion)
4. **Resource limits matter** - each session = one Excel process (~50-100MB+)
5. **LLMs must manage session lifecycle carefully** - close sessions promptly to free resources

## Problem Statement

### Current Pain Points

1. **Unintuitive Terminology**: "Begin batch" and "Commit batch" are technical terms that require explanation
2. **Cognitive Load**: LLMs must decide when to use batch mode vs. single operations
3. **Resource Leak Risk**: Forgotten commits leave Excel instances running
4. **Dual Patterns**: Tools support both `batchId` parameter and standalone operation modes (complexity!)
5. **Decision Fatigue**: LLMs waste tokens deciding "Should I use batch mode for this?"
6. **Performance Inconsistency**: Single operations are slow, batch is fast - LLM must choose correctly
7. **File Lock Race Condition (Issue #173)**: Rapid sequential non-batch calls fail because Excel disposal (2-17s) from first call hasn't completed when second call tries to open the file

### What Users Actually Think

Users and LLMs naturally think in terms of:

- **Open** a file → work with it → **Save** changes → **Close** the file
- NOT: Begin session → track GUID → commit session

This is the universal pattern for file operations across all systems.

## Proposed Solution

### High-Level Design

**Sessions are the ONLY way to work with Excel files. No exceptions.**

1. **`excel_file(action: 'open')`** - Opens a workbook, returns a `sessionId`
2. **`excel_file(action: 'save')`** - Saves changes to an open workbook
3. **`excel_file(action: 'close')`** - Closes workbook and session
4. **ALL other tools REQUIRE `sessionId`** - No standalone operation mode

**Revolutionary Change:** Remove the `batchId` optional parameter pattern entirely. Sessions are mandatory, not optional.

### Why This Works Better

| Old Pattern | New Pattern | Benefit |
|------------|-------------|---------|
| `excel_batch(action: 'begin')` | `excel_file(action: 'open')` | Matches universal file paradigm |
| Track `batchId` GUID | Track `sessionId` (still a GUID) | More intuitive name |
| `excel_batch(action: 'commit', save: true)` | `excel_file(action: 'close')` | Natural action name |
| Optional `batchId` parameter | **REQUIRED** `sessionId` parameter | No decision fatigue |
| "Should I use batch mode?" | Sessions always used | Zero cognitive load |
| Dual code paths (batch vs. single) | **Single code path only** | Simpler implementation |
| Rapid sequential calls fail (#173) | **Single Excel instance reused** | **Eliminates file lock race** |

### Terminology Changes

```
Current Term → New Term → Rationale
────────────────────────────────────────────────────────
batchId      → sessionId  → "Session" is more intuitive than "batch"
begin        → open       → Universal file operation
commit       → close      → Universal file operation (does NOT save)
batch-of-one → REMOVED    → No standalone operations anymore
Optional     → REQUIRED   → sessionId is mandatory for all operations
save param   → REMOVED    → close never saves, use explicit save action
excelPath    → REMOVED    → Session knows the file (except open/create)
```

**BREAKING CHANGE:** The optional `batchId` parameter is completely removed. Every operation on a workbook requires an active session.

## Detailed API Design

### 1. `excel_file` Tool - Updated Actions

**Current Actions:**

- `create-empty` - Create new workbook
- `close-workbook` - Emergency close (rarely used)
- `test` - Connection test

**New Actions (added):**

- **`open`** - Opens workbook, returns sessionId (replaces batch begin)
- **`save`** - Saves changes to open session
- **`close`** - Closes session WITHOUT saving (use explicit save action)

**Removed Actions:**

- None (keep create-empty, test, close-workbook for backwards compat)

### 2. API Signatures

#### Open Workbook

```csharp
[McpServerTool(Name = "excel_file")]
[Description(@"Manage Excel file lifecycle. All Excel operations require an active session.

REQUIRED WORKFLOW:
1. open - Opens workbook, returns sessionId (ALWAYS FIRST)
2. Use sessionId for ALL operations (worksheets, queries, ranges, etc.)
3. save - Saves changes (EXPLICIT action, call anytime during session)
4. close - Closes workbook and session (NEVER saves - use explicit save action)

Sessions are mandatory - there are no standalone operations.")]
public static async Task<string> ExcelFile(
    [Required]
    [Description("Action to perform")]
    FileAction action,

    [Description("Full path to Excel file - required for 'open' and 'create-empty'")]
    string? filePath = null,

    [Description("Session ID from 'open' action - required for 'save' and 'close'")]
    string? sessionId = null)
{
    return action switch
    {
        FileAction.Open => await OpenWorkbookAsync(filePath!),
        FileAction.Save => await SaveWorkbookAsync(sessionId!),
        FileAction.Close => await CloseWorkbookAsync(sessionId!),  // No save parameter - close NEVER saves
        FileAction.CreateEmpty => await CreateEmptyAsync(filePath!),
        FileAction.Test => TestConnection(),
        _ => throw new McpException($"Unknown action: {action}")
    };
}
```

#### Example Response - Open

```json
{
  "success": true,
  "sessionId": "abc-123-def-456",
  "filePath": "C:\\data\\sales.xlsx",
  "message": "Workbook opened successfully",
  "suggestedNextActions": [
    "Use sessionId='abc-123-def-456' for all operations",
    "Call excel_file(action: 'save', sessionId='...') to save changes (explicit only)",
    "Call excel_file(action: 'close', sessionId='...') when done (does NOT save)"
  ],
  "workflowHint": "Session active. Remember: close does NOT save - use explicit save action."
}
```

### 3. Other Tools - Breaking Changes

**All other tools now REQUIRE `sessionId` parameter and REMOVE `excelPath` parameter**:

```csharp
// PowerQuery example - sessionId is now REQUIRED, excelPath REMOVED
public static async Task<string> ExcelPowerQuery(
    [Required] PowerQueryAction action,
    [Required] string sessionId,  // ✅ REQUIRED (was optional batchId)
    // ... other params (excelPath REMOVED - session already knows the file)
)
```

**BREAKING CHANGE:** No more optional `batchId`. Every tool method signature changes to require `sessionId`.

**BREAKING CHANGE:** `excelPath` parameter REMOVED from all tools except `excel_file` open/create actions. The session already knows which file is open, so passing `excelPath` is redundant and creates potential for mismatches (sessionId points to fileA.xlsx, but excelPath says fileB.xlsx).

**Implementation simplification:**

- Remove `WithBatchAsync()` dual-path logic entirely
- Remove "batch-of-one" pattern
- Every tool method becomes simpler: just lookup session and use it
- No more "if sessionId provided, else create temporary batch" logic

## Implementation Strategy

### Single-Phase Breaking Refactor

**No backwards compatibility. Clean slate redesign.**

#### Step 1: Remove Old Infrastructure (1-2 days)

1. **Delete `excel_batch` tool entirely**
   - Remove `src/ExcelMcp.McpServer/Tools/BatchSessionTool.cs`
   - Remove `src/ExcelMcp.CLI/Commands/BatchCommands.cs`
   - Remove `src/ExcelMcp.McpServer/Prompts/Content/excel_batch.md`

2. **Remove dual-path logic in ExcelToolsBase**
   - Delete `WithBatchAsync()` method entirely
   - Remove "batch-of-one" pattern
   - Remove all `if (sessionId != null) { ... } else { ... }` conditionals

3. **Rename internal classes**
   - `_activeBatches` → `_activeSessions`
   - `IExcelBatch` → `IExcelSession` (interface rename)
   - `ExcelBatch` → `ExcelSession` (implementation rename)
   - `BeginBatchAsync` → `OpenSessionAsync`

#### Step 2: Add Session Lifecycle to excel_file (2-3 days)

```csharp
// Add new actions to existing excel_file tool
public enum FileAction
{
    CreateEmpty,
    Open,        // NEW - replaces batch begin
    Save,        // NEW - explicit save
    Close,       // NEW - replaces batch commit
    Test
}
```

**Implementation:**

- `OpenWorkbookAsync()` - Creates session, returns sessionId
- `SaveWorkbookAsync(sessionId)` - Saves changes
- `CloseWorkbookAsync(sessionId, save)` - Closes and disposes

#### Step 3: Update ALL Tools to Require sessionId (3-5 days)

**Before (12 tools with optional batchId):**

```csharp
string? batchId = null
```

**After (12 tools with required sessionId):**

```csharp
[Required] string sessionId
```

**Files to modify:**

- `ExcelConnectionTool.cs`
- `ExcelDataModelTool.cs`
- `ExcelFileTool.cs` (add open/save/close actions)
- `ExcelNamedRangeTool.cs`
- `ExcelPivotTableTool.cs`
- `ExcelPowerQueryTool.cs`
- `ExcelQueryTableTool.cs`
- `ExcelRangeTool.cs`
- `ExcelTableTool.cs`
- `ExcelVbaTool.cs`
- `ExcelWorksheetTool.cs`

#### Step 4: Simplify Tool Implementation (2-3 days)

**Remove complexity everywhere:**

```csharp
// OLD - Complex dual-path logic
var result = await ExcelToolsBase.WithBatchAsync(
    batchId, filePath, save: true,
    async (batch) => await commands.SomeAsync(batch, args));

// NEW - Simple direct session lookup
var session = SessionManager.GetSession(sessionId);
var result = await commands.SomeAsync(session, args);
```

**Benefits:**

- ~40% less code in each tool method
- No branching logic
- Easier to understand and maintain

#### Step 5: Update Documentation (1-2 days)

**Delete:**

- `excel_batch.md` prompt file
- All references to "batch mode" in docs
- Performance comparison sections (sessions are ALWAYS used)

**Update:**

- `excel_file.md` - Add session lifecycle patterns
- `tool_selection_guide.md` - Remove batch decision logic
- All tool descriptions - Change to "sessionId (required)"
- README - Update examples to show session workflow

## Performance & Simplification

### No More "Auto-Detection"

**Current:** LLM must decide when to use batch mode (decision fatigue)
**New:** Sessions are mandatory - no decision needed

**Key Insight:** By making sessions mandatory, we:

1. **Eliminate decision fatigue** - LLM never thinks "Should I batch?"
2. **Consistent performance** - Every operation is optimized
3. **Simpler code** - Single code path through entire system
4. **Better UX** - Open/Close workflow is intuitive

### Performance Best Practices

#### ✅ DO: Batch Operations on Same File (Single Session)

```csharp
// ✅ FAST: All operations in one session (single Excel instance reused)
var session = await CreateSessionAsync("sales.xlsx");
await GetSession(session).Execute(ctx => ctx.Book.Worksheets.Add("Q1"));  // Op 1
await GetSession(session).Execute(ctx => ctx.Book.Worksheets.Add("Q2"));  // Op 2
await GetSession(session).Execute(ctx => ctx.Book.Worksheets.Add("Q3"));  // Op 3
await SaveSessionAsync(session);
await CloseSessionAsync(session);

// Result: 1 Excel instance created/destroyed (fast)
```

#### ❌ DON'T: Open/Close Repeatedly (Multiple Excel Instances)

```csharp
// ❌ SLOW: Creates 3 separate Excel instances (5-10x overhead per open/close)
await CreateSessionAsync("sales.xlsx");
await GetSession(...).Execute(ctx => ctx.Book.Worksheets.Add("Q1"));
await CloseSessionAsync(...);

await CreateSessionAsync("sales.xlsx");  // New Excel instance!
await GetSession(...).Execute(ctx => ctx.Book.Worksheets.Add("Q2"));
await CloseSessionAsync(...);

await CreateSessionAsync("sales.xlsx");  // Yet another Excel instance!
await GetSession(...).Execute(ctx => ctx.Book.Worksheets.Add("Q3"));
await CloseSessionAsync(...);

// Result: 3 Excel instances created/destroyed (very slow)
```

**Performance Impact:** Opening Excel repeatedly is 5-10x slower than reusing a session. The session API is designed to keep Excel open for multiple operations.

#### ✅ DO: Parallel Processing of DIFFERENT Files

```csharp
// ✅ TRUE PARALLELISM: Different files = different processes
var files = new[] { "sales.xlsx", "inventory.xlsx", "customers.xlsx" };
var sessions = await Task.WhenAll(files.Select(f => CreateSessionAsync(f)));

// Process all files in parallel (3 Excel processes running simultaneously)
await Task.WhenAll(sessions.Select(async sessId => {
    var session = GetSession(sessId);
    await session.Execute(ctx => ProcessWorkbook(ctx));
    await SaveSessionAsync(sessId);
    await CloseSessionAsync(sessId);
}));

// Result: True parallelism - operations on different files don't block each other
```

#### ❌ DON'T: Try to Parallelize Operations on SAME File

```csharp
// ❌ NO BENEFIT: Operations are queued serially anyway (Excel COM limitation)
var sess = await CreateSessionAsync("data.xlsx");
await Task.WhenAll(
    GetSession(sess).Execute(ctx => ctx.Book.Worksheets.Add("Sheet1")),  // Queued
    GetSession(sess).Execute(ctx => ctx.Book.Worksheets.Add("Sheet2")),  // Queued (waits)
    GetSession(sess).Execute(ctx => ctx.Book.Worksheets.Add("Sheet3"))   // Queued (waits)
);

// Result: Operations still run serially (no speedup) - just write them sequentially for clarity
```

**Why No Benefit:** Each session has one STA thread processing operations one at a time. `Task.WhenAll` doesn't change this - they still execute serially on the Excel COM thread.

#### File Creation (Automatically Serialized)

```csharp
// ✅ This pattern works correctly - internal lock serializes calls
var tasks = Enumerable.Range(1, 10).Select(i =>
    ExcelSession.CreateNew($"report{i}.xlsx", false,
        (ctx, ct) => {
            ctx.Book.Worksheets[1].Name = $"Report {i}";
            return 0;
        }));
await Task.WhenAll(tasks);  // Executes sequentially despite Task.WhenAll!

// ✅ This explicit sequential pattern also works
for (int i = 1; i <= 10; i++)
{
    await ExcelSession.CreateNew($"report{i}.xlsx", false, ...);
}

// Result: Files created one at a time - peak memory = 1 temporary Excel instance
```

**How it works:** `ExcelSession` uses a static `SemaphoreSlim(1, 1)` to serialize all `CreateNew()` and `CreateNewAsync()` calls. Even if called via `Task.WhenAll`, they queue and execute one at a time.

**Why enforced:** Each `CreateNew()` temporarily spawns an Excel process. Without serialization, parallel calls would create N Excel processes simultaneously, causing memory exhaustion. The lock prevents this automatically.

### Code Simplification

**OLD - Complex WithBatchAsync logic:**

```csharp
public static async Task<T> WithBatchAsync<T>(
    string? batchId,
    string filePath,
    bool save,
    Func<IExcelBatch, Task<T>> action)
{
    if (!string.IsNullOrEmpty(batchId))
    {
        // Path 1: Use existing batch
        var batch = BatchSessionTool.GetBatch(batchId);
        if (batch == null) throw new McpException(...);
        if (!PathMatches(...)) throw new McpException(...);
        return await action(batch);
    }
    else
    {
        // Path 2: Create temporary "batch-of-one"
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await action(batch);
        if (save) await batch.Save();
        return result;
    }
}
```

**NEW - Simple session lookup:**

```csharp
// Every tool method becomes:
var session = SessionManager.GetSession(sessionId);  // Throws if not found
return await commands.SomeAsync(session, args);
```

**Lines of code saved:** ~200+ LOC across 12 tools

### LLM Guidance - Sessions Are Always Required

**Simplified prompt (no decisions):**

```markdown
## Excel File Operations - ALWAYS Use Sessions

**EVERY workflow follows this pattern:**
1. excel_file(action: 'open', filePath: '...') → Get sessionId
2. Perform operations (ALL require sessionId)
3. excel_file(action: 'close', sessionId: '...') → Close file

**No exceptions.** You cannot list queries, create worksheets, or read ranges without an active session.

**Single operation?** Still requires open/close:
```yaml
# Even for "just list worksheets"
1. excel_file(action: 'open', filePath: 'data.xlsx')
   → { sessionId: 'abc-123' }
2. excel_worksheet(action: 'list', sessionId: 'abc-123')
3. excel_file(action: 'close', sessionId: 'abc-123')
```

**Why?** Sessions ensure proper Excel COM lifecycle management. There are no "quick operations" - all operations are safe and optimized.

Only `excel_file(action: 'open'|'create-empty')` accepts `filePath`. All other tools use `sessionId` only and do not take a file path.

```

**Cognitive load reduced to zero:** LLM no longer decides anything about performance optimization.

## Migration Path

### No Backwards Compatibility

## Benefits Analysis

### Fixes Issue #173: File Lock Race Condition

**Problem:** In the current system, rapid sequential non-batch calls fail with file lock errors because:

1. First call creates temporary Excel instance (batch-of-one)
2. First call completes, triggers `DisposeAsync()` (2-17 seconds)
3. Second call arrives before disposal completes
4. Second call tries to open same file → **FILE LOCKED ERROR**

**How this spec eliminates the problem:**

```

Current System (Issue #173):
  Call 1: excel_range(action, NO batchId)
    → Create temp Excel → Use → Start disposal (2-17s background)
  Call 2: excel_range(action, NO batchId)
    → Try create NEW Excel → File locked! ❌

New System (Mandatory Sessions):
  Call 1: excel_file(action: 'open')
    → Create Excel instance, return sessionId
  Call 2: excel_range(action, sessionId='abc-123')
    → Reuse SAME Excel instance ✅
  Call 3: excel_range(action, sessionId='abc-123')
    → Reuse SAME Excel instance ✅
  Call 4: excel_file(action: 'close')
    → Dispose Excel once (at end)

```

**Key insight:** By requiring sessions, we eliminate the "create → dispose → create → dispose" cycle that causes the race condition. The Excel instance stays alive for the entire workflow.

**Alternative considered:** Add retry logic with exponential backoff (proposed in #173) - but this is a **workaround** for a flawed architecture. Mandatory sessions **eliminate the root cause**.

### For LLMs

| Aspect | Before | After | Improvement |
|--------|--------|-------|-------------|
| **Decision Making** | "Should I use batch mode?" | No decision - sessions always used | ✅ Zero cognitive load |
| **Parameter Naming** | `batchId` (technical) | `sessionId` (familiar) | ✅ Intuitive |
| **Workflow Clarity** | Begin→Track GUID→Commit | Open→Work→Close | ✅ Universal pattern |
| **Learning Curve** | Must understand batching | Standard file operations | ✅ No explanation needed |
| **Error Recovery** | "Did I commit?" confusion | "Is file still open?" | ✅ Natural debugging |
| **Token Efficiency** | Decide + explain batch mode | Just open/close | ✅ 50% fewer tokens |
| **Code Complexity** | Handle optional parameter | sessionId always present | ✅ Simpler reasoning |

### For Users

| Aspect | Before | After |
|--------|--------|-------|
| **Terminology** | "What's a batch?" | "Opening a file" (universal) |
| **Documentation** | Explain batch optimization | No explanation needed |
| **Error Messages** | "Batch xyz not found" | "Session xyz not found" |
| **Tool Discovery** | Find excel_batch tool | Excel_file is obvious |

### For Developers

| Aspect | Impact | Benefit |
|--------|--------|---------|
| **Code Changes** | Breaking - remove dual paths, require sessionId | ✅ ~40% less code |
| **Infrastructure** | Rename classes (Batch→Session), single path | ✅ Half the complexity |
| **Testing** | Rewrite tests for session-only workflow | ✅ Simpler test setup |
| **Backwards Compat** | None - clean break | ✅ No legacy cruft |
| **Maintenance** | Single code path to maintain | ✅ Easier debugging |
| **New Features** | Build on simpler foundation | ✅ Faster development |

## Implementation Checklist

### Core Code Changes (Breaking)

- [ ] **DELETE** `BatchSessionTool.cs` entirely
- [ ] **DELETE** `BatchCommands.cs` (CLI) entirely
- [ ] **DELETE** `excel_batch.md` prompt file
- [ ] **DELETE** `WithBatchAsync()` method in ExcelToolsBase
- [ ] **RENAME** `IExcelBatch` → `IExcelSession` (interface)
- [ ] **RENAME** `ExcelBatch` → `ExcelSession` (implementation)
- [ ] **RENAME** `_activeBatches` → `_activeSessions` in SessionManager
- [ ] **RENAME** `BeginBatchAsync` → `OpenSessionAsync` in ExcelSession
- [ ] **ADD** `FileAction.Open`, `FileAction.Save`, `FileAction.Close` enum values
- [ ] **IMPLEMENT** `OpenWorkbookAsync()`, `SaveWorkbookAsync()`, `CloseWorkbookAsync()` in ExcelFileTool
- [ ] **CHANGE** all 12 tools: `batchId` (optional) → `sessionId` (required)
- [ ] **REMOVE** `excelPath` parameter from all 11 tools (except excel_file open/create)
- [ ] **REMOVE** `save` parameter from close action in excel_file tool
- [ ] **SIMPLIFY** all tool methods: remove WithBatchAsync, direct session lookup
- [ ] **UPDATE** session to track filePath internally (for excelPath removal)

### Testing (Complete Rewrite)

- [ ] **DELETE** all batch-mode specific tests
- [ ] **REWRITE** all tool tests to use session pattern (open → operate → close)
- [ ] **ADD** session lifecycle tests (open, save, close actions)
- [ ] **ADD** error tests: operate without sessionId → clear error message
- [ ] **ADD** error tests: sessionId not found → helpful error
- [ ] **ADD** read-only workflow tests: open → read → close (no save)
- [ ] **ADD** multiple-save workflow tests: open → modify → save → modify → save → close
- [ ] **ADD** discard changes tests: open → modify → close (no save = rollback)
- [ ] **VERIFY** no performance regression (sessions were batches internally)
- [ ] **TEST** integration with MCP clients (Claude, Copilot) using new API
- [ ] **ADD** concurrency tests: verify operations within session are serial (not parallel)
- [ ] **ADD** multi-session tests: verify operations between sessions CAN run parallel
- [ ] **ADD** resource limit tests: verify 5+ concurrent sessions don't cause memory issues
- [ ] **ADD** file creation tests: verify sequential creation pattern (not parallel)

### Documentation (Complete Rewrite)

- [ ] **DELETE** `excel_batch.md` prompt file
- [ ] **DELETE** all references to "batch mode" and "when to batch"
- [ ] **REWRITE** `excel_file.md` with session lifecycle patterns
- [ ] **REWRITE** `tool_selection_guide.md` (remove batch decision logic)
- [ ] **REWRITE** README examples (all use session pattern)
- [ ] **UPDATE** all 12 tool `[Description]` attributes: "sessionId (required)"
- [ ] **ADD** session lifecycle diagram to README
- [ ] **REWRITE** `examples/` directory scripts (all use open/close)
- [ ] **ADD** migration guide: "Breaking Changes in 2.0.0"

### LLM Guidance

- [ ] Create `session_lifecycle.md` prompt with open/save/close patterns
- [ ] Update `user_request_patterns.md` with session detection hints
- [ ] Add session error recovery guidance
- [ ] Update elicitations to ask about multi-operation intent
- [ ] **ADD** concurrency model documentation: operations within session are serial
- [ ] **ADD** performance guidance: batch operations on same file, parallelize different files
- [ ] **ADD** resource limits guidance: recommend 3-5 concurrent sessions max
- [ ] **ADD** file creation guidance: always sequential, never parallel

## Edge Cases & Error Handling

### Session Not Found

**Before:**
```json
{
  "error": "Batch session 'xyz' not found. It may have already been committed..."
}
```

**After:**

```json
{
   "success": false,
   "errorMessage": "Session 'xyz' not found. The workbook may have already been closed.",
   "isError": true,
   "suggestedNextActions": [
      "Call excel_file(action: 'open', filePath: '...') to open the workbook again",
      "Check if another process closed the file"
   ]
}
```

All tools MUST return JSON for business errors:

- `success: false`
- `errorMessage`: human-readable reason
- `isError: true`
- `suggestedNextActions`: concrete next steps for the LLM

MCP exceptions (`McpException`) are reserved for protocol issues only (missing/invalid parameters, unknown actions, missing files), not business logic failures.

### Forgotten Close

**Mitigation (client-side execution model):**

1. **Process termination cleanup** - When MCP client (VS Code, Claude Desktop) closes, all Excel instances automatically close
2. **Manual process kill** - User can terminate Excel via Task Manager if needed
3. **Session listing** - Future enhancement: `excel_file(action: 'list-sessions')` to show active sessions
4. **No automatic timeout** - Client-side execution means no server-side cleanup needed

**Why this works:**

- MCP server runs on user's machine (not remote server)
- Excel process lifetime tied to MCP client process lifetime
- User has full control via OS process management

### File Locking

No change - same Excel COM behavior. Sessions don't change locking semantics.

### Multiple Workbooks / Sessions

**You can open multiple files simultaneously, but understand the concurrency model:**

1. Each file gets its own session (and Excel process)
2. Operations **between** files run in parallel (separate processes)
3. Operations **within** each file remain serial (Excel COM limitation - see CRITICAL section above)

**Example:**

```csharp
// Open 3 files (3 Excel processes created)
var sessA = await CreateSessionAsync("A.xlsx");  // Excel.Application process 1
var sessB = await CreateSessionAsync("B.xlsx");  // Excel.Application process 2
var sessC = await CreateSessionAsync("C.xlsx");  // Excel.Application process 3

// These 3 operations run in TRUE parallel (different processes)
await Task.WhenAll(
    GetSession(sessA).Execute(ctx => ctx.Book.Worksheets.Add("Sheet1")),  // Process 1
    GetSession(sessB).Execute(ctx => ctx.Book.Worksheets.Add("Sheet1")),  // Process 2
    GetSession(sessC).Execute(ctx => ctx.Book.Worksheets.Add("Sheet1"))   // Process 3
);  // ✅ True parallelism - different Excel processes

// But operations on SAME file are SERIAL (queued)
await GetSession(sessA).Execute(ctx => ctx.Book.Worksheets.Add("Sheet2"));  // Queued operation 1
await GetSession(sessA).Execute(ctx => ctx.Book.Worksheets.Add("Sheet3"));  // Queued operation 2 (waits for 1)
```

**Resource Limits & Best Practices:**

- Each session = one Excel process (~50-100MB+ memory)
- **Recommendation:** Limit to 3-5 concurrent sessions for typical desktop machines
- LLMs should close sessions promptly to free resources
- Monitor system resources when processing many files

**File Locking:**

- SessionManager prevents opening same file twice: `File 'X.xlsx' is already open in another session`
- This matches Excel UI behavior (cannot open same file in multiple windows)
- Attempting to open an already-open file throws `InvalidOperationException`

**LLM Guidance:**

LLMs should track the correct `sessionId` for each workbook and close sessions when each logical workflow completes. For bulk file processing, consider sequential processing to limit resource usage:

```
# Processing many files - sequential approach (resource-friendly)
for each file:
  1. open → sessionId
  2. perform operations (all serial within session)
  3. save
  4. close (frees Excel process immediately)

# OR parallel approach for small batches (faster but memory-intensive)
1. open files 1-5 → get 5 sessionIds (5 Excel processes)
2. process all 5 in parallel
3. close all 5
4. repeat for files 6-10
```

## Success Metrics

### Quantitative

- **Reduced token usage**: Session guidance ~40% shorter than batch guidance
- **Fewer errors**: Track "session not found" vs. "batch not found" rates
- **Adoption rate**: % of multi-operation workflows using sessions

### Qualitative

- **LLM feedback**: Do Claude/Copilot naturally use open/close without prompting?
- **User confusion**: Reduced questions about "what's a batch?" in docs/issues
- **Code clarity**: Naming matches intent in tool descriptions


## Design Decisions (Resolved)

### 1. Should `open` action fail if workbook already open in Excel UI?

**Decision:** Yes, fail immediately with clear error
**Rationale:** Excel COM limitation - we can't safely work with UI-open files
**Implementation:** Existing file lock detection works correctly

### 2. Should `save` be implicit on `close` by default?

**Decision:** No, close NEVER saves - explicit save action only
**Rationale:**

- **Explicit is better than implicit** - No surprise saves
- **LLM clarity** - "save" action = save, "close" action = close (no overlap)
- **Read-only workflows** - Just open → read → close (no save needed)
- **Multiple saves** - Save multiple times during session, close at end
- **Predictable behavior** - close always does same thing (cleanup only)

**Implementation:** Remove `save` parameter from close entirely. Users call `excel_file(action: 'save')` explicitly when needed.

### 3. Should sessions timeout automatically after inactivity?

**Decision:** No automatic timeout - rely on process lifetime
**Rationale:**

- **Client-side execution** - MCP server runs on user's machine, not remote server
- **Process lifetime** - When user closes MCP client (VS Code, Claude Desktop), process terminates and Excel closes
- **Manual control** - User can kill Excel process via Task Manager if needed
- **Simpler implementation** - No background timers, no timeout logic
- **No false positives** - No "session timed out" errors during long-running operations

**Implementation:** No timeout logic. Sessions persist until explicitly closed or process terminates.

### 4. What are the save semantics for different workflows?

**Decision:** Explicit save action only, close never saves

**Workflows supported:**

1. **Read-only workflow:**

   ```
   open → read operations → close (no save needed)
   ```

   Use case: List queries, view data, check connections

2. **Single save workflow:**

   ```
   open → modify operations → save → close
   ```

   Use case: Standard edit workflow

3. **Multiple save workflow:**

   ```
   open → modify → save → modify → save → modify → save → close
   ```

   Use case: Incremental changes, checkpoints, complex multi-step operations

4. **Discard changes workflow:**

   ```
   open → modify operations → close (no save = changes discarded)
   ```

   Use case: Experimental changes, testing, rollback

**Rationale:**

- **Explicit control** - User decides when to persist changes
- **Read-only support** - No save parameter needed anywhere
- **Flexibility** - Save 0, 1, or N times during session
- **Predictability** - close always does same thing (cleanup)

### 5. Should we keep `excel_batch` as deprecated alias?

**Decision:** No, complete removal
**Rationale:**

- Maintaining alias adds complexity
- Breaking change anyway, might as well be clean
- Forces users to adopt new pattern completely
- No confusion from "two ways to do same thing"

### 6. Should CLI and MCP Server both use same session API?

**Decision:** Yes, unified API everywhere
**Rationale:**

- CLI and MCP Server share Core/ComInterop
- Consistent experience across interfaces
- Same documentation applies to both
- No mode-specific quirks

## Timeline

**Estimated effort:** 2-3 weeks (one developer)

**Week 1: Delete & Rename (Breaking Changes)**

- Day 1-2: Delete batch infrastructure, rename classes
- Day 3-4: Add session lifecycle to excel_file tool
- Day 5: Update 3-4 tools to require sessionId

**Week 2: Tool Updates & Testing**

- Day 1-3: Update remaining 8-9 tools to require sessionId
- Day 4: Simplify all tool implementations (remove WithBatchAsync)
- Day 5: Rewrite core tests for session-only pattern

**Week 3: Integration & Documentation**

- Day 1-2: Integration tests with MCP clients
- Day 3: Rewrite all documentation and examples
- Day 4: Migration guide, release notes, breaking change announcements
- Day 5: Beta release testing

**Release Timeline:**

- Week 4: Version 2.0.0-beta (breaking changes, early adopters)
- Week 6: Version 2.0.0 stable (after beta feedback)
- Month 6: Version 1.x end-of-life (final security patch)

## Appendix: Example Workflows

### Before (Batch API)

```
LLM: I'll create 3 worksheets using batch mode for performance.

1. excel_batch(action: 'begin', filePath: 'sales.xlsx')
   → { batchId: 'abc-123' }

2. excel_worksheet(action: 'create', excelPath: 'sales.xlsx',
                   sheetName: 'Q1', batchId: 'abc-123')

3. excel_worksheet(action: 'create', excelPath: 'sales.xlsx',
                   sheetName: 'Q2', batchId: 'abc-123')

4. excel_worksheet(action: 'create', excelPath: 'sales.xlsx',
                   sheetName: 'Q3', batchId: 'abc-123')

5. excel_batch(action: 'commit', batchId: 'abc-123', save: true)
   → { success: true }
```

### After (Session API)

```
LLM: I'll open the workbook and create 3 worksheets.

1. excel_file(action: 'open', filePath: 'sales.xlsx')
   → { sessionId: 'abc-123' }

2. excel_worksheet(action: 'create',
                   sheetName: 'Q1', sessionId: 'abc-123')

3. excel_worksheet(action: 'create',
                   sheetName: 'Q2', sessionId: 'abc-123')

4. excel_worksheet(action: 'create',
                   sheetName: 'Q3', sessionId: 'abc-123')

5. excel_file(action: 'close', sessionId: 'abc-123')
   → { success: true }
```

**Differences:**

- "Begin batch" → "Open file" (natural language)
- "Commit batch" → "Close file" (universal action)
- `batchId` (optional) → `sessionId` (required)
- No more decision about "should I batch?"
- **Close never saves** - explicit save action only
- **excelPath removed** from all operations except open (session knows the file)
- Simpler: 5 calls vs 5 calls, but with intuitive naming

### Read-Only Workflow Example

```
LLM: I'll check which Power Queries are in this workbook.

1. excel_file(action: 'open', filePath: 'sales.xlsx')
   → { sessionId: 'abc-123' }

2. excel_powerquery(action: 'list', sessionId: 'abc-123')
   → { queries: ['SalesData', 'CustomerInfo', 'ProductCatalog'] }

3. excel_file(action: 'close', sessionId: 'abc-123')
   → { success: true }
```

**Key points:**

- No save action needed (read-only operation)
- Close doesn't save (no changes made)
- Simple: open → read → close

### Multiple-Save Workflow Example

```
LLM: I'll create multiple queries with checkpoints after each one.

1. excel_file(action: 'open', filePath: 'sales.xlsx')
   → { sessionId: 'abc-123' }

2. excel_powerquery(action: 'import', sessionId: 'abc-123',
                    queryName: 'SalesData', mCodeFile: 'sales.m')
   → { success: true }

3. excel_file(action: 'save', sessionId: 'abc-123')
   → { success: true }  // Checkpoint 1

4. excel_powerquery(action: 'import', sessionId: 'abc-123',
                    queryName: 'CustomerInfo', mCodeFile: 'customers.m')
   → { success: true }

5. excel_file(action: 'save', sessionId: 'abc-123')
   → { success: true }  // Checkpoint 2

6. excel_powerquery(action: 'import', sessionId: 'abc-123',
                    queryName: 'ProductCatalog', mCodeFile: 'products.m')
   → { success: true }

7. excel_file(action: 'save', sessionId: 'abc-123')
   → { success: true }  // Final checkpoint

8. excel_file(action: 'close', sessionId: 'abc-123')
   → { success: true }
```

**Key points:**

- Multiple explicit save actions during session
- Each save creates a checkpoint (changes persisted)
- Close at end does NOT save (last save already persisted everything)
- Incremental persistence reduces risk of data loss

## Conclusion

This redesign achieves the ultimate goal: **Eliminate all cognitive load from LLMs by making sessions the only way to work with Excel files**. The Open/Save/Close pattern is:

1. **Universal** - Every developer/LLM knows file lifecycle (no explanation needed)
2. **Mandatory** - No decisions about batching, sessions are always used
3. **Simple** - Single code path, 40% less code to maintain
4. **Performant** - Same Excel COM optimization (sessions are batches internally)
5. **Breaking** - Clean slate, no backwards compatibility baggage
6. **Extensible** - Future optimizations (connection pooling, caching) build on simpler foundation
7. **Bug-fixing** - Eliminates Issue #173 file lock race condition by design
8. **Realistic** - Acknowledges and documents Excel COM threading limitations (single-threaded STA model)

### Key Achievements

**For LLMs:**

- ✅ Zero decision fatigue (no "should I batch?" questions)
- ✅ 50% fewer tokens for workflows (no batch mode explanations)
- ✅ Intuitive API (open/close is universal)
- ✅ Clear concurrency model (operations within session = serial, between sessions = parallel)

**For Developers:**

- ✅ 40% less code (remove dual paths)
- ✅ Simpler testing (single pattern)
- ✅ Easier maintenance (one way to do things)
- ✅ Fixes Issue #173 (eliminates file lock race condition at architectural level)
- ✅ Honest documentation (acknowledges COM limitations, not hidden complexity)

**For Users:**

- ✅ Consistent performance (always optimized)
- ✅ Clear errors ("session not found" is obvious)
- ✅ Predictable behavior (explicit lifecycle)
- ✅ No more file lock errors from rapid sequential operations
- ✅ Realistic expectations (understand that operations within a file are serial due to Excel COM)

### Critical Technical Constraints

**This spec acknowledges and documents fundamental Excel COM limitations:**

1. **Single-threaded nature** - Each session runs on one STA thread with serial operation queue
2. **No within-session parallelism** - Multiple operations on same file execute serially (COM requirement)
3. **Between-session parallelism possible** - Different files = different Excel processes = true parallelism
4. **Resource constraints** - Each session = ~50-100MB+ memory; recommend 3-5 concurrent sessions max
5. **File creation must be sequential** - Parallel creation causes resource exhaustion

These are **not implementation deficiencies** - they are inherent Excel COM API constraints that apply to ANY Excel automation solution (VBA, .NET interop, Python xlwings, etc.). By documenting them clearly, we set correct expectations and guide users toward performant patterns.

### Breaking Change Justification

**Why break backwards compatibility?**

1. **Current API is fundamentally flawed** - Optional batching creates decision fatigue
2. **Gradual migration would take years** - Dual paths would persist indefinitely
3. **Clean break is clearer** - Users update once vs. confused by deprecated patterns
4. **Version 2.0 is the right time** - Major version signals breaking changes
5. **Migration is straightforward** - Wrap operations in open/close (mechanical change)
6. **Fixes critical bug at architectural level** - Issue #173 file lock race condition eliminated by design (retry logic is just a workaround)

**Recommendation:** ✅ **Approve for implementation in Version 2.0.0 (breaking release)**

This is not just a rename - it's a fundamental simplification that makes ExcelMcp significantly easier for LLMs to use correctly while eliminating an entire class of file locking bugs.

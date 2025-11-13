# Session API Redesign Specification

**Version:** 1.0  
**Status:** Draft  
**Date:** 2025-01-13  
**Author:** Development Team

## Executive Summary

This specification proposes a **breaking redesign** of ExcelMcp's session API to use intuitive **Open/Save/Close** semantics exclusively. The goal is to eliminate the "batch" concept entirely and remove all cognitive load from LLMs - every operation works through sessions, always. No backwards compatibility, no dual patterns, no decisions about when to batch.

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
        if (save) await batch.SaveAsync();
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

You can have multiple sessions open at once (one per workbook):

1. `excel_file(action: 'open', filePath: 'A.xlsx')` → `sessionId: 'sess-A'`
2. `excel_file(action: 'open', filePath: 'B.xlsx')` → `sessionId: 'sess-B'`
3. Use `sess-A` for operations on `A.xlsx`, and `sess-B` for operations on `B.xlsx`.
4. Close each when done: `excel_file(action: 'close', sessionId: 'sess-A')`, then `'sess-B'`.

LLMs should track and reuse the correct `sessionId` for each workbook, and close sessions when each logical workflow completes.

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

### Key Achievements

**For LLMs:**

- ✅ Zero decision fatigue (no "should I batch?" questions)
- ✅ 50% fewer tokens for workflows (no batch mode explanations)
- ✅ Intuitive API (open/close is universal)

**For Developers:**

- ✅ 40% less code (remove dual paths)
- ✅ Simpler testing (single pattern)
- ✅ Easier maintenance (one way to do things)
- ✅ Fixes Issue #173 (eliminates file lock race condition at architectural level)

**For Users:**

- ✅ Consistent performance (always optimized)
- ✅ Clear errors ("session not found" is obvious)
- ✅ Predictable behavior (explicit lifecycle)
- ✅ No more file lock errors from rapid sequential operations

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

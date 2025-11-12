# File Handle Refactoring Specification

> **Replace batch sessions with explicit file handle management**

## Current State Problems

### 1. Batch Concept is Confusing
- LLMs struggle to understand when to use batch mode
- Users don't know if they need `batchId` or not
- Documentation is complex (elicitations, prompts, workflow hints)
- Two modes of operation: batch vs batch-of-one (implicit)

### 2. Architectural Issues
- `ExcelSession` and `ExcelBatch` mix concerns
- Batch disposal takes 2-17 seconds (causes Issue #173)
- File locking issues with rapid sequential operations
- Complex lifetime management (STA threads, COM cleanup)

### 3. API Inconsistency
- Some operations work on "file path", some on "batch"
- MCP Server has `excel_batch` tool (extra concept to learn)
- CLI has batch commands (inconsistent with other commands)

## Proposed Solution: FilePath-Based Commands with Handle Caching

### Core Concept

**FilePath as Parameter** = All commands take `filePath` as first parameter. FileHandleManager caches handles by path and reuses them automatically.

```csharp
// Core API - Commands take filePath
public interface IRangeCommands
{
    Task<OperationResult> GetValuesAsync(string filePath, string sheetName, string rangeAddress);
    Task<OperationResult> SetValuesAsync(string filePath, string sheetName, string rangeAddress, object[,] values);
}

// NEW: Explicit workbook lifecycle commands (optional - for explicit control)
public interface IWorkbookCommands
{
    Task<OperationResult> CreateAsync(string filePath);     // Create new workbook
    Task<OperationResult> SaveAsync(string filePath);       // Explicitly save workbook
    Task<OperationResult> CloseAsync(string filePath);      // Explicitly close workbook (releases handle)
}
```

**Why filePath parameter:**
- MCP Server and CLI are **stateless** - each tool call/command is independent
- No "active workbook" persists across separate HTTP requests or CLI invocations
- FilePath is the natural key for handle caching

**Note:** `IWorkbookCommands` is a NEW interface. Existing `IFileCommands` remains untouched (no build breakage). After Phase 5 cleanup, `IFileCommands` will be removed entirely.

### How It Works (Human/LLM Workflow)

**External API (MCP/CLI):** Users pass file paths - each call is independent (stateless)
**Internal Implementation:** FileHandleManager caches handles by path, reuses automatically

```
LLM Call 1: "Get data from report.xlsx"
  → excel_range(excelPath: "report.xlsx", action: "get-values", sheet: "Sheet1", range: "A1:D10")
  → Internally: FileHandleManager opens Excel → caches handle by path → returns data

LLM Call 2: "Add data to Sheet2 in report.xlsx"
  → excel_range(excelPath: "report.xlsx", action: "set-values", sheet: "Sheet2", ...)
  → Internally: FileHandleManager REUSES handle from Call 1 (NO LOCK!)

LLM Call 3: "Close report.xlsx"
  → excel_file(action: "close", excelPath: "report.xlsx")
  → Internally: FileHandleManager releases handle → disposes Excel
```

**Key insight:** Each MCP tool call or CLI command is stateless, but FileHandleManager maintains state internally (caches handles by absolute path).

**Automatic cleanup:** Handles close after inactivity timeout (e.g., 5 minutes) if not explicitly closed.

**Multi-file workflows:**
```
LLM: "Copy data from source.xlsx to target.xlsx"
  → excel_range(excelPath: "source.xlsx", action: "get-values", ...) → data
  → excel_range(excelPath: "target.xlsx", action: "set-values", ..., data)
  → Both files cached, both reused on subsequent calls
  → excel_file(action: "close", excelPath: "target.xlsx")
  → excel_file(action: "close", excelPath: "source.xlsx")
```

## Architecture Design

### 1. Active Workbook Context (Internal)

```csharp
// Thread-safe active workbook tracking using AsyncLocal
internal static class ActiveWorkbook
{
    private static readonly AsyncLocal<FileHandle?> _current = new();
    
    internal static FileHandle Current 
    { 
        get => _current.Value ?? throw new InvalidOperationException("No active workbook. Call OpenAsync() or CreateAsync() first.");
        set => _current.Value = value;
    }
    
    internal static bool HasActive => _current.Value != null;
}
```

**Why AsyncLocal?**
- Thread-safe for async/await code
- Each async call chain has its own "current" workbook
- Tests can run in parallel without interference

### 2. File Handle Structure

```csharp
internal sealed class FileHandle
{
    public string Id { get; }              // Unique identifier (GUID)
    public string FilePath { get; }        // Absolute path to workbook
    public DateTime OpenedAt { get; }      // When opened
    public bool IsClosed { get; internal set; }  // Lifecycle state
    
    internal FileHandle(string id, string filePath)
    {
        Id = id;
        FilePath = Path.GetFullPath(filePath);
        OpenedAt = DateTime.UtcNow;
        IsClosed = false;
    }
}
```

### 3. File Handle Manager (Internal Cache)

```csharp
// Internal to ExcelMcp.Core - caches handles by absolute file path
internal sealed class FileHandleManager : IDisposable
{
    // Singleton instance - one per process (CLI or MCP Server)
    private static readonly Lazy<FileHandleManager> _instance = new(() => new FileHandleManager());
    public static FileHandleManager Instance => _instance.Value;
    
    // Cache handles by absolute file path
    private readonly ConcurrentDictionary<string, ExcelWorkbookHandle> _handles = new();
    private readonly SemaphoreSlim _lock = new(1, 1);
    
    /// <summary>
    /// Opens or retrieves cached handle for file path. Thread-safe.
    /// </summary>
    internal async Task<ExcelWorkbookHandle> OpenOrGetAsync(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);
        
        await _lock.WaitAsync();
        try
        {
            // Reuse existing handle if already open
            if (_handles.TryGetValue(absolutePath, out var existing))
            {
                existing.UpdateLastAccess();  // Reset inactivity timeout
                return existing;
            }
            
            // Create new handle
            var newHandle = await ExcelWorkbookHandle.CreateAsync(absolutePath);
            _handles[absolutePath] = newHandle;
            return newHandle;
        }
        finally
        {
            _lock.Release();
        }
    }
    
    /// <summary>
    /// Gets handle for file path (must already be open).
    /// </summary>
    internal ExcelWorkbookHandle GetHandle(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);
        
        if (_handles.TryGetValue(absolutePath, out var handle))
            return handle;
            
        throw new InvalidOperationException($"File not open: {filePath}");
    }
    
    /// <summary>
    /// Explicitly close handle for file path. Removes from cache and disposes.
    /// </summary>
    internal async Task CloseAsync(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);
        
        await _lock.WaitAsync();
        try
        {
            if (_handles.TryRemove(absolutePath, out var handle))
            {
                await handle.DisposeAsync();
            }
        }
        finally
        {
            _lock.Release();
        }
    }
    
    /// <summary>
    /// Save specific file (must already be open).
    /// </summary>
    internal async Task SaveAsync(string filePath)
    {
        var handle = GetHandle(filePath);  // Throws if not open
        await handle.SaveAsync();
    }
    
    /// <summary>
    /// Background cleanup: Close handles inactive for > timeout.
    /// Runs periodically (e.g., every minute).
    /// </summary>
    internal async Task CleanupInactiveHandlesAsync(TimeSpan inactivityTimeout)
    {
        var now = DateTime.UtcNow;
        var toRemove = _handles
            .Where(kvp => (now - kvp.Value.LastAccess) > inactivityTimeout)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var path in toRemove)
        {
            await CloseAsync(path);
        }
    }
    
    public void Dispose()
    {
        foreach (var handle in _handles.Values)
        {
            handle.Dispose();
        }
        _handles.Clear();
        _lock.Dispose();
    }
}

/// <summary>
/// Wraps Excel COM objects for a single workbook.
/// Tracks last access time for automatic cleanup.
/// </summary>
internal sealed class ExcelWorkbookHandle : IAsyncDisposable
{
    private dynamic? _application;
    private dynamic? _workbook;
    
    public string FilePath { get; }
    public DateTime LastAccess { get; private set; }
    public dynamic Application => _application ?? throw new ObjectDisposedException(nameof(ExcelWorkbookHandle));
    public dynamic Workbook => _workbook ?? throw new ObjectDisposedException(nameof(ExcelWorkbookHandle));

    private ExcelWorkbookHandle(string filePath)
    {
        FilePath = Path.GetFullPath(filePath);
        LastAccess = DateTime.UtcNow;
    }

    public static async Task<ExcelWorkbookHandle> CreateAsync(string filePath)
    {
        var handle = new ExcelWorkbookHandle(filePath);
        await Task.Run(() =>
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) throw new InvalidOperationException("Excel not installed");
            
            handle._application = Activator.CreateInstance(excelType);
            handle._application.Visible = false;
            handle._application.DisplayAlerts = false;
            handle._workbook = handle._application.Workbooks.Open(filePath);
        });
        return handle;
    }

    public void UpdateLastAccess() => LastAccess = DateTime.UtcNow;

    public async Task SaveAsync()
    {
        await Task.Run(() => _workbook?.Save());
        UpdateLastAccess();
    }

    public async ValueTask DisposeAsync()
    {
        await Task.Run(() =>
        {
            try
            {
                _workbook?.Close(false);
                ComUtilities.Release(ref _workbook);
                
                _application?.Quit();
                ComUtilities.Release(ref _application);
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { /* Suppress disposal errors */ }
        });
    }

    public void Dispose()
    {
        DisposeAsync().AsTask().GetAwaiter().GetResult();
    }
}
```

### 3. How This Solves Issue #173

**Issue #173**: Sequential operations on same file caused locks because LLMs didn't use batch mode

**Before (Batch Mode - LLM must opt-in):**
```
Operation 1: Open Excel → work → close (2-17s disposal) → LOCKS FILE
Operation 2: Can't open - file locked!
```

**After (FilePath-Based with Handle Caching - Automatic):**
```
LLM Call 1: excel_range(excelPath: "report.xlsx", action: "get-values", ...)
  → Internal: FileHandleManager opens Excel → caches by path → returns data

LLM Call 2: excel_range(excelPath: "report.xlsx", action: "set-values", ...)
  → Internal: FileHandleManager REUSES cached handle (NO NEW OPEN, NO LOCK!)

LLM Call 3 (much later or never): excel_file(action: "close", excelPath: "report.xlsx")
  → Internal: FileHandleManager closes and removes handle
  OR: Background cleanup closes after 5 minutes inactivity
```

**Key insight**: 
- **External API**: Stateless - each call passes filePath
- **Internal Implementation**: Stateful cache - handles keyed by absolute path
- **Result**: Multiple calls to same file reuse handle automatically (no opt-in needed!)
- **Cleanup**: Explicit close OR timeout-based background disposal

### 4. Core Commands Pattern

```csharp
// Example: RangeCommands
public class RangeCommands : IRangeCommands
{
    public async Task<OperationResult> GetValuesAsync(
        string filePath,
        string sheetName, 
        string rangeAddress)
    {
        // Open or get cached handle
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        
        return await Task.Run(() =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = handle.Workbook.Worksheets[sheetName];
                range = sheet.Range[rangeAddress];
                object[,] values = range.Value2;
                
                return new OperationResult
                {
                    Success = true,
                    Data = new { Values = values }
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }
}
```

**Pattern**: Every command method takes `filePath` as first parameter, uses FileHandleManager to get/reuse handle.

## Migration Strategy: Sequential End-to-End Conversion

**Approach:** Convert each command interface end-to-end (Core → Tests → MCP → CLI) in separate commits. All commits stay in ONE pull request for Issue #173.

**Why Sequential End-to-End Commits:**
- ✅ Logical progression (review one command at a time in commit history)
- ✅ Each command is fully validated before moving to next commit
- ✅ Incremental progress (can pause/resume at any command boundary)
- ✅ Batch API remains functional for unconverted commands
- ✅ Reduces risk (smaller blast radius per commit)
- ✅ ONE PR for Issue #173 (all commits together)

### Phase 0: Foundation (One-Time Setup - First Commit)

### Phase 0: Foundation (One-Time Setup - First Commit)

**Goal:** Create shared infrastructure used by all commands

1. Add `ExcelWorkbookHandle` class to `ExcelMcp.ComInterop/Session/` (wraps COM objects)
2. Add `FileHandleManager` to `ExcelMcp.ComInterop/Session/` (singleton cache)
3. **Create NEW interface** `IWorkbookCommands` in `ExcelMcp.Core/Commands/`:
   - `Task CreateAsync(string filePath)`
   - `Task OpenAsync(string filePath)` - explicit open (optional - manager auto-opens)
   - `Task SaveAsync(string filePath)`
   - `Task CloseAsync(string filePath)`
4. **Create NEW class** `WorkbookCommands` implementing `IWorkbookCommands`
5. Write unit tests for `FileHandleManager` (handle caching, reuse, cleanup)
6. Add background cleanup task (timeout-based disposal)

**Why NEW interface `IWorkbookCommands`:**
- Clean separation (workbook lifecycle vs feature operations)
- `IFileCommands` deleted entirely in final cleanup (no gradual migration)

**Deliverable:** ONE commit with FileHandleManager + IWorkbookCommands + tests

---

### Per-Command Migration Pattern (One Commit Per Command)

For each command interface (IRangeCommands, IPowerQueryCommands, ISheetCommands, etc.), create **ONE commit** with all 4 steps:

#### Step 1: Core Commands (Add FilePath-Based Methods)

1. Add filePath-based methods to interface:
   ```csharp
   // NEW: FilePath-based API (filePath is first parameter!)
   Task<OperationResult> GetValuesAsync(string filePath, string sheetName, string rangeAddress);
   
   // OLD: Batch API (keep for now - deleted in final cleanup)
   Task<OperationResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress);
   ```

2. Implement filePath-based methods in Commands class:
   ```csharp
   public async Task<OperationResult> GetValuesAsync(string filePath, string sheetName, string rangeAddress)
   {
       var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
       
       return await Task.Run(() =>
       {
           dynamic? sheet = null;
           dynamic? range = null;
           try
           {
               sheet = handle.Workbook.Worksheets[sheetName];
               range = sheet.Range[rangeAddress];
               // ... work with range ...
           }
           finally
           {
               ComUtilities.Release(ref range!);
               ComUtilities.Release(ref sheet!);
           }
       });
   }
   ```

3. Build passes (0 warnings)

**Deliverable:** Core Commands updated for ONE command interface

#### Step 2: Tests (Convert to FilePath-Based)

1. Update test class to inject `IWorkbookCommands` (for explicit close)
2. Convert all tests for this command interface:
   ```csharp
   // OLD (batch-based)
   await using var batch = await ExcelSession.BeginBatchAsync(testFile);
   var result = await _commands.GetValuesAsync(batch, "Sheet1", "A1:D10");
   
   // NEW (filePath-based)
   var result = await _commands.GetValuesAsync(testFile, "Sheet1", "A1:D10");
   await _workbookCommands.SaveAsync(testFile);  // Explicit save if needed
   await _workbookCommands.CloseAsync(testFile);  // Explicit close
   ```

3. Run tests for this command interface - all pass
4. Verify handle reuse works (multiple operations on same file)

**Deliverable:** Tests updated and passing for ONE command interface

#### Step 3: MCP Server (Update Tool)

1. Update MCP tool for this command:
   ```csharp
   // OLD (batch-based)
   await using var batch = await ExcelSession.BeginBatchAsync(excelPath);
   var result = await commands.GetValuesAsync(batch, sheet, range);
   await batch.SaveAsync();
   
   // NEW (filePath-based)
   var result = await commands.GetValuesAsync(excelPath, sheet, range);  // Pass filePath!
   await workbookCommands.SaveAsync(excelPath);  // Explicit save
   ```

2. Update tool `[Description]` if needed (usually no changes - tools already accept excelPath)
3. Run MCP Server tests for this tool - all pass

**Deliverable:** MCP tool updated and tested for ONE command interface

#### Step 4: CLI (Update Commands)

1. Update CLI command handlers for this command:
   ```csharp
   // OLD (batch-based)
   await using var batch = await ExcelSession.BeginBatchAsync(filePath);
   var result = await _commands.GetValuesAsync(batch, sheet, range);
   await batch.SaveAsync();
   
   // NEW (filePath-based)
   var result = await _commands.GetValuesAsync(filePath, sheet, range);  // Pass filePath!
   await _workbookCommands.SaveAsync(filePath);  // Explicit save
   ```

2. No changes to CLI command signatures - users still pass file paths
3. Run CLI tests for this command - all pass

**Deliverable:** CLI commands updated and tested for ONE command interface

---

### Command Migration Order (Suggested)

**Start simple, build confidence:**

1. **IWorksheetCommands** (simple, no complex COM interactions)
2. **INamedRangeCommands** (simple, good for testing handle caching)
3. **IRangeCommands** (most used, validates performance)
4. **ITableCommands** (moderate complexity)
5. **IPowerQueryCommands** (heavy operations, validates timeout handling)
6. **IConnectionCommands** (validates multi-file scenarios)
7. **IDataModelCommands** (complex COM, good stress test)
8. **IPivotTableCommands** (complex COM)
9. **IQueryTableCommands** (moderate complexity)
10. **IVbaCommands** (special handling for .xlsm)

**Each command = One commit with all 4 steps complete**

---

### Final Cleanup Phase (After All Commands Converted - Final Commit)

**Goal:** Remove batch infrastructure completely in ONE final commit

1. Verify all commands migrated (all interfaces use filePath-based API)
2. Delete batch infrastructure in one clean commit:
   - `ExcelSession.cs`
   - `ExcelBatch.cs`
   - `IExcelBatch.cs`
   - `IFileCommands.cs` interface (replaced by `IWorkbookCommands`)
   - `FileCommands.cs` implementation (replaced by `WorkbookCommands`)
   - Batch-based method overloads from all `I*Commands` interfaces
   - `excel_batch` MCP tool (no longer needed)
   - Batch CLI commands (`batch-begin`, `batch-commit`, `batch-discard`)
3. Remove batch-related documentation
4. Update CHANGELOG

**Why this works:**
- All consumers already migrated (one command at a time)
- No external consumers (Core is internal library)
- No `[Obsolete]` warnings to clean up
- Single atomic cleanup commit (easy to review)

**Deliverable:** Clean codebase with zero batch references

## Benefits

### 1. Conceptual Simplicity
- ✅ Matches stateless HTTP/CLI model: each call passes filePath
- ✅ Matches LLM workflow: "work with this file" (just pass path)
- ✅ No batch mode confusion - handle reuse is automatic!
- ✅ No prompts/elicitations needed (LLMs already pass file paths)
- ✅ Clear lifecycle: pass path → work → optionally close (or auto-cleanup after timeout)

### 2. API Simplicity  
- ✅ **FilePath as first parameter** on all command methods (explicit, clear)
- ✅ Method signatures: `GetValuesAsync(filePath, sheet, range)` - stateless!
- ✅ Less cognitive load for developers and LLMs (just pass the path)
- ✅ Multi-file scenarios just work (each call specifies which file)

### 3. Performance
- ✅ Excel stays open across operations (automatic handle caching by path)
- ✅ No 2-17 second disposal between operations
- ✅ Fixes Issue #173 (sequential file lock errors)
- ✅ Explicit close OR background timeout cleanup

### 4. Architecture
- ✅ Clear separation: FileHandleManager manages cache internally, consumers just pass filePath
- ✅ Thread-safe via ConcurrentDictionary and SemaphoreSlim
- ✅ No STA thread complexity exposed to consumers
- ✅ Simpler testing (pass filePath, no batch setup)
- ✅ Easier to understand and maintain
- ✅ Stateless-compatible: Works perfectly with HTTP/CLI request-response model

### 5. Compatibility
- ✅ Consistent: all operations use filePath parameter
- ✅ Explicit: LLM decides when to save/close OR auto-cleanup after timeout
- ✅ Flexible: multiple files work naturally (each call specifies which file)
- ✅ Debuggable: filePath in logs/errors

## Comparison: Before vs After

### Before (Batch Mode)

**MCP Server:**
```
LLM: I want to add data to Sheet1 and Sheet2 in report.xlsx

Tool calls:
1. excel_batch(action: "begin", filePath: "report.xlsx") 
   → {batchId: "batch_abc"}
2. excel_range(batchId: "batch_abc", action: "set-values", sheet: "Sheet1", ...)
3. excel_range(batchId: "batch_abc", action: "set-values", sheet: "Sheet2", ...)
4. excel_batch(action: "commit", batchId: "batch_abc", save: true)
```

**Problems:**
- 4 tool calls
- Batch concept to explain
- batchId to track
- When to commit?
- LLMs didn't use it → sequential file locks (Issue #173)

### After (FilePath-Based with Handle Caching)

**MCP Server:**
```
LLM: I want to add data to Sheet1 and Sheet2 in report.xlsx

Tool calls:
1. excel_range(excelPath: "report.xlsx", action: "set-values", sheet: "Sheet1", ...)
   → Internal: FileHandleManager opens Excel → caches handle by path
2. excel_range(excelPath: "report.xlsx", action: "set-values", sheet: "Sheet2", ...)
   → Internal: FileHandleManager REUSES cached handle (NO LOCK!)
3. (optional) excel_file(action: "close", excelPath: "report.xlsx")
   → Internal: FileHandleManager releases handle
```

**Benefits:**
- 2 tool calls (3 if explicit close)
- Natural workflow (just pass file path)
- No new concepts (LLMs already pass file paths!)
- Handle reuse automatic (fixes Issue #173)
- Cleanup automatic (timeout-based) OR explicit

## Edge Cases & Considerations

### 1. Multiple Files Open

**Scenario:** LLM wants to copy data between two workbooks

```
1. excel_range(excelPath: "source.xlsx", action: "get-values", ...) → data
   → Internal: Opens source.xlsx, caches handle
2. excel_range(excelPath: "target.xlsx", action: "set-values", ..., data)
   → Internal: Opens target.xlsx, caches handle (both cached now!)
3. excel_file(action: "close", excelPath: "target.xlsx")
4. excel_file(action: "close", excelPath: "source.xlsx")
```

✅ Supported naturally - FileHandleManager caches multiple files by path

### 2. Forgot to Close

**Problem:** LLM passes filePath but never explicitly closes

**Solution:** 
- Background cleanup task runs periodically (e.g., every minute)
- Closes handles inactive for > timeout (e.g., 5 minutes)
- Add `list_open_files()` diagnostic tool (shows cached file paths)
- Graceful shutdown disposes all handles

**Scenario:** LLM wants to discard changes

```
excel_file(action: "close", excelPath: "report.xlsx")
```

Internal: FileHandleManager closes without saving (default behavior)

For explicit save:
```
excel_file(action: "save", excelPath: "report.xlsx")
excel_file(action: "close", excelPath: "report.xlsx")
```

✅ Explicit control over save behavior

### 4. File Already Open

**Problem:** Try to work on file that's already cached

**Solution:** FileHandleManager automatically reuses cached handle:
```csharp
public async Task<ExcelWorkbookHandle> OpenOrGetAsync(string filePath)
{
    string absolutePath = Path.GetFullPath(filePath);
    
    // Check cache first - REUSE if exists!
    if (_handles.TryGetValue(absolutePath, out var existing))
    {
        existing.UpdateLastAccess();  // Reset timeout
        return existing;  // ✅ Reuse cached handle (fixes Issue #173!)
    }
    
    // Open new handle if not cached
    var newHandle = await ExcelWorkbookHandle.CreateAsync(absolutePath);
    _handles[absolutePath] = newHandle;
    return newHandle;
}
```

✅ Handle reuse automatic - no error, no confusion!

### 5. Testing Strategy

**Unit Tests:** Easy - pass filePath
```csharp
[Fact]
public async Task GetValues_ValidFilePath_ReturnsData()
{
    // Create test file
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(...);
    
    // Test operation - just pass filePath!
    var result = await _commands.GetValuesAsync(testFile, "Sheet1", "A1:D10");
    Assert.True(result.Success);
    
    // Cleanup (optional - timeout will handle it)
    await _workbookCommands.CloseAsync(testFile);
}
```

**Integration Tests:** 
```csharp
[Fact]
public async Task FullWorkflow_PassFilePathEverywhere_Succeeds()
{
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(...);
    
    // Work - handle reuse automatic!
    await _rangeCommands.SetValuesAsync(testFile, "Sheet1", "A1", data);
    await _rangeCommands.SetValuesAsync(testFile, "Sheet2", "B2", data2);
    
    // Save & Close
    await _workbookCommands.SaveAsync(testFile);
    await _workbookCommands.CloseAsync(testFile);
    
    // Verify persisted
    var result = await _rangeCommands.GetValuesAsync(testFile, "Sheet1", "A1");
    Assert.Equal(expectedValue, result.Data.Values[0,0]);
    await _workbookCommands.CloseAsync(testFile);
}
```

**Migration Strategy:**
- Sequential end-to-end conversion (Core → Tests → MCP → CLI)
- One commit per command interface
- All commits in ONE pull request for Issue #173
- No dual test maintenance needed

## Implementation Checklist

### Phase 0: Foundation (First Commit)
- [ ] Create `ExcelWorkbookHandle` class in `ExcelMcp.ComInterop/Session/`
- [ ] Create `FileHandleManager` singleton in `ExcelMcp.ComInterop/Session/`
- [ ] Create `IWorkbookCommands` interface in `ExcelMcp.Core/Commands/`
- [ ] Implement `WorkbookCommands` class
- [ ] Write unit tests for `FileHandleManager` (handle caching, reuse, cleanup)
- [ ] Add background cleanup task (timeout-based disposal)
- [ ] **PR:** Infrastructure foundation

### Per-Command Conversion (One Commit Per Command)

**For each command in migration order, create ONE commit with all 4 steps:**

#### Command: [IWorksheetCommands / INamedRangeCommands / IRangeCommands / etc.]

- [ ] **Step 1 - Core:**
  - [ ] Add filePath-based methods to interface (keep batch overloads)
  - [ ] Implement filePath-based methods in Commands class
  - [ ] Build passes (0 warnings)

- [ ] **Step 2 - Tests:**
  - [ ] Update test class to inject `IWorkbookCommands`
  - [ ] Convert all tests to filePath-based API
  - [ ] Run tests for this command - all pass
  - [ ] Verify handle reuse works

- [ ] **Step 3 - MCP Server:**
  - [ ] Update MCP tool to pass filePath to Core Commands
  - [ ] Remove batch session logic from tool
  - [ ] Update tool `[Description]` if needed
  - [ ] Run MCP Server tests - all pass

- [ ] **Step 4 - CLI:**
  - [ ] Update CLI command handlers to pass filePath
  - [ ] Remove batch session logic
  - [ ] Run CLI tests - all pass

- [ ] **COMMIT:** "Convert [CommandName] to filePath-based API (Core+Tests+MCP+CLI)"

### Final Cleanup Phase (Final Commit After All Commands)

- [ ] **Delete Batch Infrastructure:**
  - [ ] `ExcelSession.cs`
  - [ ] `ExcelBatch.cs`
  - [ ] `IExcelBatch.cs`
  - [ ] `IFileCommands.cs` and `FileCommands.cs`
  - [ ] Batch method overloads from all `I*Commands` interfaces
  - [ ] `excel_batch` MCP tool
  - [ ] Batch CLI commands

- [ ] **Documentation Cleanup:**
  - [ ] Remove all `SuggestedNextActions` from MCP Server responses
  - [ ] Update all tool `[Description]` attributes (no batch guidance)
  - [ ] Delete ALL prompt files (LLMs pass filePaths naturally)
  - [ ] Update README.md (remove batch mode, show automatic handle reuse)
  - [ ] Update MCP Server README
  - [ ] Update CLI README

- [ ] **CHANGELOG:**
  - [ ] Document breaking change (batch API removed)
  - [ ] Document new filePath-based API with automatic handle caching

- [ ] **Final Verification:**
  - [ ] Run full test suite (all filePath-based)
  - [ ] Grep search for batch references (except docs/CHANGELOG)
  - [ ] Build passes (0 warnings)

- [ ] **PR:** Final cleanup - remove all batch infrastructure

## Risk Assessment
  - [ ] Delete ALL prompt files (LLMs already pass file paths)
  - [ ] Clean up elicitations directory
- [ ] **CLI:**
  - [ ] Update all command help text to remove batch references
  - [ ] Update `--help` output for all commands
  - [ ] Remove batch examples from CLI README
- [ ] **Documentation:**
  - [ ] Update README.md (remove batch mode, emphasize automatic handle reuse)
  - [ ] Update MCP Server README (new tool descriptions, no batch concept)
  - [ ] Update examples (filePath patterns, show multi-file scenarios)
  - [ ] Add migration guide (batch → filePath)
  - [ ] Update architecture documentation
- [ ] **Testing:**
  - [ ] Run full test suite (all filePath-based)
  - [ ] Verify no batch infrastructure references remain
  
- [ ] **COMMIT:** "Final cleanup: Remove batch infrastructure"

---

## Pull Request Structure

**One PR for Issue #173 with ~12 commits:**

1. **Commit 1:** Phase 0 - Foundation (FileHandleManager + IWorkbookCommands + tests)
2. **Commit 2:** Convert IWorksheetCommands (Core+Tests+MCP+CLI)
3. **Commit 3:** Convert INamedRangeCommands (Core+Tests+MCP+CLI)
4. **Commit 4:** Convert IRangeCommands (Core+Tests+MCP+CLI)
5. **Commit 5:** Convert ITableCommands (Core+Tests+MCP+CLI)
6. **Commit 6:** Convert IPowerQueryCommands (Core+Tests+MCP+CLI)
7. **Commit 7:** Convert IConnectionCommands (Core+Tests+MCP+CLI)
8. **Commit 8:** Convert IDataModelCommands (Core+Tests+MCP+CLI)
9. **Commit 9:** Convert IPivotTableCommands (Core+Tests+MCP+CLI)
10. **Commit 10:** Convert IQueryTableCommands (Core+Tests+MCP+CLI)
11. **Commit 11:** Convert IVbaCommands (Core+Tests+MCP+CLI)
12. **Commit 12:** Final cleanup - Remove batch infrastructure

**Benefits:**
- ✅ Logical progression in commit history (review one command at a time)
- ✅ Easy to bisect if issues found later
- ✅ Can cherry-pick specific command conversions if needed
- ✅ Single PR keeps all related work together
- ✅ Closes Issue #173 with complete solution

---

## Risk Assessment

### Low Risk
- ✅ Sequential end-to-end conversion (one commit per command)
- ✅ Each command fully tested before moving to next commit
- ✅ Incremental progress with pause/resume capability at commit boundaries
- ✅ Batch API functional for unconverted commands
- ✅ Clean removal (no deprecation warnings)
- ✅ FilePath-based API is stateless-compatible (HTTP/CLI natural fit)

### Medium Risk
- ⚠️ Large refactor (all Commands classes)
- ⚠️ Multiple layers (Core, CLI, MCP Server)
- ⚠️ No external consumer testing (internal library only)

**Mitigation:**
- **Sequential conversion:** One commit per command interface (logical progression)
- **Batch API functional:** Unconverted commands still work during migration
- **Comprehensive testing:** Each command fully tested before moving on
- **Can pause/resume:** Stop after any commit, resume later

### High Risk
- ❌ None identified

## Success Criteria

1. ✅ LLM can work with files by just passing filePath (no batch concept, no explicit handle management)
2. ✅ No file lock errors on sequential operations (Issue #173 fixed via automatic handle reuse)
3. ✅ All tests use filePath-based API (100% coverage)
4. ✅ Documentation is clearer (simpler - just pass filePath!)
5. ✅ Performance equal or better (Excel stays open, handles cached)
6. ✅ **Per-command validation:** Each command's 4 steps complete before moving on
7. ✅ **Final cleanup completeness:**
   - Zero references to "batch" in code (except historical docs/CHANGELOG)
   - Zero `SuggestedNextActions` in MCP Server responses
   - All tool descriptions updated (no batch guidance)
   - All CLI help text updated (no batch references)
   - `IFileCommands` and `FileCommands` deleted (replaced by `IWorkbookCommands`)
   - **All tests use filePath-based API** (zero batch-based tests remain)

## Questions for Iteration

1. **Handle lifetime:** Auto-expire after timeout (e.g., 5 minutes inactivity)?
2. **CLI behavior:** Auto-close after single operation, or keep open until explicit close?
3. **Multiple files:** Limit on concurrent open handles? (probably not needed)
4. **Error recovery:** What if Close fails? Orphaned Excel process?
5. **Phase 5 timing:** Remove batch immediately after Phase 4, or delay for additional testing?
6. **Background cleanup:** How often? What timeout? Per-process or global?

## Related Issues

- Fixes #173: Sequential file lock errors (root cause eliminated via automatic handle caching by path)
- Eliminates need for prompts/elicitations (LLMs already pass file paths naturally!)
- Improves LLM workflow clarity (stateless HTTP/CLI model with invisible caching)
- Matches universal programming pattern (pass filePath, system handles optimization)

---

**Next Steps:**
1. Review and iterate on this spec
2. Create GitHub issue/project for tracking
3. Start Phase 1 implementation
4. Get early feedback on file handle API

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

## Proposed Solution: Active Workbook Pattern

### Core Concept

**Active Workbook** = The currently open Excel workbook that all operations work on (similar to Excel's UI)

```csharp
// Core API - NEW workbook lifecycle management
public interface IWorkbookCommands
{
    Task CreateAsync(string filePath);     // Create and set as active
    Task OpenAsync(string filePath);       // Open and set as active
    Task SaveAsync();                      // Save active workbook
    Task CloseAsync();                     // Close active workbook
    Task CloseAsync(string filePath);      // Close specific workbook (for multi-file scenarios)
}

// All operations work on active workbook (NO handle parameter!)
public interface IRangeCommands
{
    Task<OperationResult> GetValuesAsync(string sheetName, string rangeAddress);
    Task<OperationResult> SetValuesAsync(string sheetName, string rangeAddress, object[,] values);
}
```

**Note:** `IWorkbookCommands` is a NEW interface. Existing `IFileCommands` remains untouched (no build breakage). After Phase 5 cleanup, `IFileCommands` will be removed entirely.

### How It Works (Human/LLM Workflow)

**External API (MCP/CLI):** Users continue to pass file paths - no change
**Internal Implementation:** Active workbook managed behind the scenes

```
LLM: "Get data from report.xlsx"

Internally:
1. MCP Server opens file → sets as active workbook
2. Calls Core Command → GetValuesAsync(sheet, range)  [no file parameter!]
3. Saves and closes → CloseAsync()
4. Returns result to LLM

User workflow is simple: open → work → close
```

**For multiple operations on same file:**
```
LLM: "Add data to Sheet1 in report.xlsx"
  → Open report.xlsx (sets as active)
  → SetValuesAsync("Sheet1", "A1", data)

LLM: "Now add another sheet"
  → CreateSheetAsync("NewSheet")  [works on active workbook - report.xlsx]
  
LLM: "Save and close"
  → SaveAsync(), CloseAsync()
```

**For multi-file workflows (rare):**
```
LLM: "Copy data from source.xlsx to target.xlsx"
  → OpenAsync("source.xlsx")
  → data = GetValuesAsync("Sheet1", "A1:D10")
  → OpenAsync("target.xlsx")  [switches active workbook]
  → SetValuesAsync("Sheet1", "A1", data)
  → CloseAsync("target.xlsx")
  → CloseAsync("source.xlsx")
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

### 3. File Handle Manager (Internal)

```csharp
// Internal to ExcelMcp.Core - not exposed to consumers
internal sealed class FileHandleManager
{
    // Key by absolute file path - reuses handles for same file
    private static readonly ConcurrentDictionary<string, (FileHandle Handle, dynamic Excel, dynamic Workbook)> _handlesByPath = new();
    
    internal static FileHandle Create(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);
        
        // Check if already open
        if (_handlesByPath.TryGetValue(absolutePath, out var existing))
        {
            ActiveWorkbook.Current = existing.Handle;
            return existing.Handle;  // Reuse existing handle
        }
        
        var handle = new FileHandle(Guid.NewGuid().ToString(), absolutePath);
        var (excel, workbook) = OpenExcelInstance(absolutePath, createNew: true);
        _handlesByPath[absolutePath] = (handle, excel, workbook);
        ActiveWorkbook.Current = handle;
        return handle;
    }
    
    internal static FileHandle Open(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);
        
        // Check if already open
        if (_handlesByPath.TryGetValue(absolutePath, out var existing))
        {
            ActiveWorkbook.Current = existing.Handle;
            return existing.Handle;  // Reuse existing handle (solves Issue #173!)
        }
        
        var handle = new FileHandle(Guid.NewGuid().ToString(), absolutePath);
        var (excel, workbook) = OpenExcelInstance(absolutePath, createNew: false);
        _handlesByPath[absolutePath] = (handle, excel, workbook);
        ActiveWorkbook.Current = handle;
        return handle;
    }
    
    internal static (dynamic Excel, dynamic Workbook) GetActiveWorkbook()
    {
        var handle = ActiveWorkbook.Current;
        
        // Find by file path
        if (!_handlesByPath.TryGetValue(handle.FilePath, out var entry))
            throw new InvalidOperationException($"Active workbook not found: {handle.FilePath}");
            
        return (entry.Excel, entry.Workbook);
    }
    
    internal static void Close(string? filePath = null)
    {
        // If filePath specified, close that file; otherwise close active workbook
        string targetPath = filePath != null 
            ? Path.GetFullPath(filePath) 
            : ActiveWorkbook.Current.FilePath;
        
        // Remove by file path
        if (!_handlesByPath.TryRemove(targetPath, out var entry))
            return;
            
        try
        {
            entry.Workbook.Close(SaveChanges: false);
            entry.Excel.Quit();
            
            // COM cleanup
            ComUtilities.Release(ref entry.Workbook!);
            ComUtilities.Release(ref entry.Excel!);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        finally
        {
            entry.Handle.IsClosed = true;
            
            // Clear active workbook if this was it
            if (ActiveWorkbook.HasActive && ActiveWorkbook.Current.FilePath == targetPath)
                ActiveWorkbook.Current = null!;
        }
    }
    
    internal static void Save()
    {
        var (_, workbook) = GetActiveWorkbook();
        workbook.Save();
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

**After (Active Workbook Pattern - Automatic):**
```
Operation 1: OpenAsync("report.xlsx") → returns existing handle if open → sets as active → work
Operation 2: GetValuesAsync(...) → uses active workbook → work (NO LOCK!)
Operation 3: CloseAsync() → only now does disposal happen
```

**Key insight**: File handles are cached by path. Multiple `OpenAsync()` calls for same file return the same handle and set it as active. This makes "batch mode" the default behavior without LLMs needing to understand it.

### 4. Core Commands Pattern

```csharp
// Example: RangeCommands
public class RangeCommands : IRangeCommands
{
    public async Task<OperationResult> GetValuesAsync(
        string sheetName, 
        string rangeAddress)
    {
        return await Task.Run(() =>
        {
            // Get active workbook - NO handle parameter!
            var (excel, workbook) = FileHandleManager.GetActiveWorkbook();
            
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = workbook.Worksheets[sheetName];
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

## Migration Strategy

### Phase 1: Add Active Workbook API (Parallel to Batch)

**Goal:** Introduce active workbook pattern without breaking existing code

1. Add `FileHandle` class to `ExcelMcp.Core/Models/` (internal)
2. Add `FileHandleManager` to `ExcelMcp.Core/Session/` (internal)
3. Add `ActiveWorkbook` context to `ExcelMcp.Core/Session/` (internal)
4. **Create NEW interface** `IWorkbookCommands` in `ExcelMcp.Core/Commands/`:
   - `Task CreateAsync(string filePath)`
   - `Task OpenAsync(string filePath)`
   - `Task SaveAsync()`
   - `Task CloseAsync()`
   - `Task CloseAsync(string filePath)` - for multi-file scenarios
5. **Create NEW class** `WorkbookCommands` implementing `IWorkbookCommands`
6. Update all `I*Commands` interfaces to remove handle parameter:
   ```csharp
   // NEW: Active workbook API (no handle parameter!)
   Task<OperationResult> GetValuesAsync(string sheetName, string rangeAddress);
   
   // OLD: Batch API (keep during development, remove in Phase 5)
   Task<OperationResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress);
   ```
7. Keep old batch-based methods for backward compatibility (NO obsolete attribute - removes cleanly later)

**Why NEW interface instead of modifying `IFileCommands`:**
- No build breakage (existing `IFileCommands` unchanged)
- Clean separation (old file operations vs new workbook lifecycle)
- `IFileCommands` deleted entirely in Phase 5 (no gradual migration)

### Phase 2: Update MCP Server

**Goal:** Replace `IFileCommands` usage with `IWorkbookCommands` internally

**External API:** No change - tools still accept `excelPath` parameter
**Internal Implementation:** Replace batch session with active workbook management

1. Update `ExcelToolsBase.WithBatchAsync()` → `WithActiveWorkbookAsync()`:
   ```csharp
   // OLD (batch-based)
   await using var batch = await ExcelSession.BeginBatchAsync(excelPath);
   var result = await commands.GetValuesAsync(batch, sheet, range);
   await batch.SaveAsync();
   
   // NEW (active workbook pattern)
   await workbookCommands.OpenAsync(excelPath);  // Sets as active, reuses if already open
   try {
       var result = await commands.GetValuesAsync(sheet, range);  // No handle parameter!
       await workbookCommands.SaveAsync();
   } finally {
       await workbookCommands.CloseAsync();
   }
   ```

2. Remove `excel_batch` tool (no longer needed)
3. Update all tool implementations to use new pattern
4. **No changes to tool signatures** - users still pass file paths

### Phase 3: Update CLI

**Goal:** Replace `IFileCommands` usage with `IWorkbookCommands` internally

**External API:** No change - commands still accept `--file` parameter
**Internal Implementation:** Replace batch session with active workbook management

1. Update all command handlers:
   ```csharp
   // OLD (batch-based)
   await using var batch = await ExcelSession.BeginBatchAsync(filePath);
   var result = await _commands.GetValuesAsync(batch, sheet, range);
   await batch.SaveAsync();
   
   // NEW (active workbook pattern)  
   await _workbookCommands.OpenAsync(filePath);  // Sets as active
   try {
       var result = await _commands.GetValuesAsync(sheet, range);  // No handle parameter!
       await _workbookCommands.SaveAsync();
   } finally {
       await _workbookCommands.CloseAsync();
   }
   ```

2. Remove batch CLI commands (`batch-begin`, `batch-commit`, etc.)
3. **No changes to command signatures** - users still pass `--file` flag

**CLI auto-manages active workbook:**
```bash
# CLI opens, sets as active, executes, saves, closes automatically
excelmcp range get-values --file report.xlsx --sheet Sheet1 --range A1:D10
```

### Phase 4: Update Tests

**Goal:** Migrate all tests to active workbook API

1. Replace batch test pattern:
   ```csharp
   // OLD (batch-based)
   await using var batch = await ExcelSession.BeginBatchAsync(testFile);
   var result = await _commands.GetValuesAsync(batch, "Sheet1", "A1:D10");
   
   // NEW (active workbook pattern)
   await _workbookCommands.OpenAsync(testFile);
   try {
       var result = await _commands.GetValuesAsync("Sheet1", "A1:D10");  // No handle parameter!
   } finally {
       await _workbookCommands.CloseAsync();
   }
   ```

2. Update all test classes to inject `IWorkbookCommands`
3. Replace `ExcelSession.BeginBatchAsync()` with `OpenAsync()`
4. No more `await using var batch` - explicit open/close instead
5. Benefits:
   - Tests exercise actual production code paths
   - Faster disposal (immediate close, not 2-17 second delay)
   - Clearer test intent (open → work → close)
   - Simpler API (no handle parameters in command calls)

### Phase 5: Remove Batch Infrastructure

**Goal:** Clean removal of obsolete batch code

1. Verify all consumers migrated (Phases 1-4: Core APIs, MCP Server, CLI, tests all use file handles)
2. Delete batch infrastructure in one clean commit:
   - `ExcelSession.cs`
   - `ExcelBatch.cs`
   - `IExcelBatch.cs`
   - `IFileCommands.cs` interface (replaced by `IWorkbookCommands`)
   - `FileCommands.cs` implementation (replaced by `WorkbookCommands`)
   - Batch-based method overloads from all `I*Commands` interfaces
   - `excel_batch` MCP tool (no longer needed)
   - Batch CLI commands (no longer needed)
3. Remove batch-related documentation
4. Update CHANGELOG

**Why clean removal works:**
- No external consumers (Core is internal library)
- All internal consumers migrated in Phases 2-4 (MCP Server, CLI, tests)
- No `[Obsolete]` warnings to clean up
- Single atomic change (easy to review)

## Benefits

### 1. Conceptual Simplicity
- ✅ Matches human mental model: "open file, work, save, close"
- ✅ Matches Excel's UI model: one active workbook at a time
- ✅ No batch mode confusion - it's just how the system works
- ✅ No prompts/elicitations needed (LLMs already know this pattern)
- ✅ Clear lifecycle: create/open → work → save/close

### 2. API Simplicity  
- ✅ **No handle parameters** on ~160 command methods
- ✅ Simpler method signatures: `GetValuesAsync(sheet, range)` not `GetValuesAsync(handle, sheet, range)`
- ✅ Less cognitive load for developers and LLMs
- ✅ Still supports rare multi-file scenarios (via explicit file path in CloseAsync)

### 3. Performance
- ✅ Excel stays open across operations (automatic handle reuse)
- ✅ No 2-17 second disposal between operations
- ✅ Fixes Issue #173 (sequential file lock errors)
- ✅ LLM controls when to close (vs automatic disposal)

### 4. Architecture
- ✅ Clear separation: Core manages active workbook, consumers don't see handles
- ✅ Thread-safe via AsyncLocal (each async chain has its own active workbook)
- ✅ No STA thread complexity exposed to consumers
- ✅ Simpler testing (pass handle, no batch setup)
- ✅ Easier to understand and maintain

### 4. API Design
- ✅ Consistent: all operations use `FileHandle`
- ✅ Explicit: LLM decides when to save/close
- ✅ Flexible: multiple files open simultaneously
- ✅ Debuggable: handle ID in logs/errors

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

### After (File Handle)

**MCP Server:**
```
LLM: I want to add data to Sheet1 and Sheet2 in report.xlsx

Tool calls:
1. open_file(filePath: "report.xlsx") 
   → {handle: "abc123"}
2. excel_range(handle: "abc123", action: "set-values", sheet: "Sheet1", ...)
3. excel_range(handle: "abc123", action: "set-values", sheet: "Sheet2", ...)
4. save_file(handle: "abc123")
5. close_file(handle: "abc123")
```

**Benefits:**
- 5 tool calls (same as batch + explicit close)
- Intuitive workflow (open→work→save→close)
- No new concepts
- Clear lifecycle

## Edge Cases & Considerations

### 1. Multiple Files Open

**Scenario:** LLM wants to copy data between two workbooks

```
1. open_file("source.xlsx") → handle1
2. open_file("target.xlsx") → handle2
3. excel_range(handle1, "get-values", ...) → data
4. excel_range(handle2, "set-values", ..., data)
5. save_file(handle2)
6. close_file(handle1)
7. close_file(handle2)
```

✅ Supported naturally with file handles

### 2. Forgot to Close

**Problem:** LLM opens file but never closes it

**Solution:** 
- MCP Server could track open handles per session
- Implement timeout (auto-close after 5 minutes of inactivity)
- Add `list_open_files()` tool for diagnostics
- Add health check that warns about abandoned handles

### 3. Close Without Save

**Scenario:** LLM wants to discard changes

```
close_file(handle: "abc123", save: false)
```

✅ Explicit control over save behavior

### 4. File Already Open

**Problem:** Try to open file that's already open

**Solution:**
```csharp
public async Task<FileHandle> OpenAsync(string filePath)
{
    // Check if file already open
    var existingHandle = FileHandleManager.FindByPath(filePath);
    if (existingHandle != null)
    {
        return new OperationResult
        {
            Success = false,
            ErrorMessage = $"File already open with handle {existingHandle.Id}. Use that handle or close it first."
        };
    }
    
    // Normal open logic...
}
```

### 5. Testing Strategy

**Unit Tests:** Easy - pass mock FileHandle
```csharp
[Fact]
public async Task GetValues_ValidHandle_ReturnsData()
{
    // Create file and get handle
    var handle = await _fileCommands.CreateAsync(_testFile);
    
    // Test operation
    var result = await _commands.GetValuesAsync(handle, "Sheet1", "A1:D10");
    Assert.True(result.Success);
    
    // Cleanup
    await _fileCommands.CloseAsync(handle);
}
```

**Integration Tests:** 
```csharp
[Fact]
public async Task FullWorkflow_OpenWorkSaveClose_Succeeds()
{
    // Open
    var handle = await _fileCommands.OpenAsync(_testFile);
    
    // Work
    await _rangeCommands.SetValuesAsync(handle, "Sheet1", "A1", data);
    
    // Save & Close
    await _fileCommands.SaveAsync(handle);
    await _fileCommands.CloseAsync(handle);
    
    // Verify persisted
    var verifyHandle = await _fileCommands.OpenAsync(_testFile);
    var result = await _rangeCommands.GetValuesAsync(verifyHandle, "Sheet1", "A1");
    Assert.Equal(expectedValue, result.Data.Values[0,0]);
    await _fileCommands.CloseAsync(verifyHandle);
}
```

**Migration Strategy:**
- Phase 1: Migrate all tests to file handle API immediately
- Delete batch-based tests as soon as file handle implementation is complete
- No dual test maintenance needed

## Implementation Checklist

### Phase 1: Core Infrastructure (Week 1)
- [ ] Create `FileHandle` class
- [ ] Create `FileHandleManager` (internal)
- [ ] Add file commands to `IFileCommands`
- [ ] Implement `FileCommands` class
- [ ] Add file handle overloads to all `I*Commands` interfaces (keep batch overloads)
- [ ] Update all Commands classes with file handle implementations
- [ ] Write unit tests for `FileHandleManager`
- [ ] Write integration tests for file handle lifecycle
- [ ] **Migrate ALL tests to file handle API immediately**
- [ ] **Delete batch-based tests** (no backward compat verification needed)

### Phase 2: MCP Server (Week 2)
- [ ] Create `FileHandleTool.cs` (create, open, save, close)
- [ ] Add `handle` parameter to all existing tools
- [ ] Update tool descriptions to show file handle API (no workflow guidance needed)
- [ ] Write MCP Server tests (file handle workflow)
- [ ] Keep batch tool functional (for now)
- [ ] **Note:** Don't remove `SuggestedNextActions` or prompts yet - Phase 5 cleanup

### Phase 3: CLI (Week 2)
- [ ] Add file commands (`file create`, `file open`, etc.)
- [ ] Add `--handle` flag to all commands
- [ ] Update help text (remove batch references)
- [ ] Update command `--help` output to show file handle usage
- [ ] Write CLI tests (file handle workflow)
- [ ] Keep batch commands functional (for now)
- [ ] **Note:** Full help text cleanup in Phase 5

### Phase 4: Documentation (Week 3)
- [ ] Update README.md (remove batch mode, show file handle)
- [ ] Update MCP Server README (new tool descriptions)
- [ ] Update examples (file handle patterns)
- [ ] Add migration guide (batch → file handle)
- [ ] Update architecture docs

### Phase 5: Cleanup (Week 3-4)
- [ ] **Core Commands:**
  - [ ] Remove batch method overloads from all `I*Commands` interfaces
  - [ ] Delete `ExcelSession.cs`, `ExcelBatch.cs`, `IExcelBatch.cs`
  - [ ] Remove unused methods from `IFileCommands` (keep only Create, Open, Save, Close)
- [ ] **MCP Server:**
  - [ ] Delete batch tool (`BatchSessionTool.cs`)
  - [ ] Remove all `SuggestedNextActions` from all tool responses (182 instances)
  - [ ] Update all tool `[Description]` attributes to remove batch mode guidance
  - [ ] Delete ALL prompt files (LLMs already understand open→work→close pattern)
  - [ ] Update elicitations directory - LLMs already understand open→work→close pattern
- [ ] **CLI:**
  - [ ] Delete batch commands (`batch-begin`, `batch-commit`, `batch-discard`)
  - [ ] Update all command help text to remove batch references
  - [ ] Update `--help` output for all commands
  - [ ] Remove batch examples from CLI README
- [ ] **Documentation:**
  - [ ] Remove batch documentation from all README files
  - [ ] Remove batch examples
  - [ ] Update architecture documentation
- [ ] **Testing:**
  - [ ] Run full test suite (all file handle-based)
  - [ ] Verify no batch-related code remains (grep search)
- [ ] **CHANGELOG:**
  - [ ] Document breaking change (batch API removed)
  - [ ] Document new file handle API

## Risk Assessment

### Low Risk
- ✅ Backward compatible during development (Phases 1-3)
- ✅ Incremental rollout (phases)
- ✅ Tests validate behavior
- ✅ Clean removal (no deprecation warnings)

### Medium Risk
- ⚠️ Large refactor (all Commands classes)
- ⚠️ Multiple layers (Core, CLI, MCP Server)
- ⚠️ No external consumer testing (internal library only)

**Mitigation:**
- Keep batch API functional until Phase 5 (consumers migrate incrementally)
- Comprehensive test coverage for file handle API
- All tests migrated to file handle in Phase 1

### High Risk
- ❌ None identified

## Success Criteria

1. ✅ LLM can open→work→save→close without batch concept
2. ✅ No file lock errors on sequential operations (Issue #173 fixed)
3. ✅ All tests use file handle API (100% coverage)
4. ✅ Documentation is clearer (simpler)
5. ✅ Performance equal or better (Excel stays open)
7. ✅ **Phase 5 completeness checks:**
   - Zero references to "batch" in code (except historical docs/CHANGELOG)
   - Zero `SuggestedNextActions` in MCP Server responses
   - All tool descriptions updated (no batch guidance)
   - All CLI help text updated (no batch references)
   - Only 4 methods in `IFileCommands` (Create, Open, Save, Close)
   - **All tests use file handle API** (zero batch-based tests remain)

## Questions for Iteration

1. **Handle lifetime:** Should handles auto-expire after timeout?
2. **Handle format:** GUID string or structured (e.g., `file_abc123`)?
3. **CLI auto-management:** Should CLI auto-open/close for single operations?
4. **Multiple files:** Limit on concurrent open handles?
5. **Error recovery:** What if Close fails? Orphaned Excel process?
6. **Phase 5 timing:** Remove batch immediately after Phase 4, or delay for additional testing?

## Related Issues

- Fixes #173: Sequential file lock errors (root cause eliminated)
- Eliminates need for prompts/elicitations (LLMs understand file handles natively)
- Improves LLM workflow clarity (matches universal programming pattern)

---

**Next Steps:**
1. Review and iterate on this spec
2. Create GitHub issue/project for tracking
3. Start Phase 1 implementation
4. Get early feedback on file handle API

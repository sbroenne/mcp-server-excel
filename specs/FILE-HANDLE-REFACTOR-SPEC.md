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

## Proposed Solution: File Handles

### Core Concept

**File Handle** = An opaque identifier representing an open Excel workbook

```csharp
// Core API
public interface IFileCommands
{
    Task<FileHandle> CreateAsync(string filePath);
    Task<FileHandle> OpenAsync(string filePath);
    Task SaveAsync(FileHandle handle);
    Task CloseAsync(FileHandle handle);
}

// All operations use file handle
public interface IRangeCommands
{
    Task<OperationResult> GetValuesAsync(FileHandle handle, string sheetName, string rangeAddress);
    Task<OperationResult> SetValuesAsync(FileHandle handle, string sheetName, string rangeAddress, object[,] values);
}
```

**Note:** After cleanup (Phase 5), `IFileCommands` will only have these 4 methods. All other file operations (like `Test`) will be removed if not needed for the new architecture.

### How It Works (Human/LLM Workflow)

```
1. LLM: "Create a new workbook" → create_file → FileHandle{id: "abc123"}
2. LLM: "Add data to Sheet1"   → excel_range(handle: "abc123", action: "set-values", ...)
3. LLM: "Save the file"        → save_file(handle: "abc123")
4. LLM: "Close the file"       → close_file(handle: "abc123")
```

**Or for existing files:**
```
1. LLM: "Open report.xlsx"     → open_file → FileHandle{id: "xyz789"}
2. LLM: "Get data from Sheet2" → excel_range(handle: "xyz789", action: "get-values", ...)
3. LLM: "Close without saving" → close_file(handle: "xyz789", save: false)
```

## Architecture Design

### 1. File Handle Structure

```csharp
public sealed class FileHandle
{
    public string Id { get; }              // Unique identifier (GUID)
    public string FilePath { get; }        // Absolute path to workbook
    public DateTime OpenedAt { get; }      // When opened
    public bool IsClosed { get; internal set; }  // Lifecycle state
    
    internal FileHandle(string id, string filePath)
    {
        Id = id;
        FilePath = filePath;
        OpenedAt = DateTime.UtcNow;
        IsClosed = false;
    }
}
```

### 2. File Handle Manager (Internal)

```csharp
// Internal to ExcelMcp.Core - not exposed to consumers
internal sealed class FileHandleManager
{
    private static readonly ConcurrentDictionary<string, (dynamic Excel, dynamic Workbook)> _handles = new();
    
    internal static FileHandle Create(string filePath)
    {
        var handle = new FileHandle(Guid.NewGuid().ToString(), filePath);
        var (excel, workbook) = OpenExcelInstance(filePath, createNew: true);
        _handles[handle.Id] = (excel, workbook);
        return handle;
    }
    
    internal static FileHandle Open(string filePath)
    {
        var handle = new FileHandle(Guid.NewGuid().ToString(), filePath);
        var (excel, workbook) = OpenExcelInstance(filePath, createNew: false);
        _handles[handle.Id] = (excel, workbook);
        return handle;
    }
    
    internal static (dynamic Excel, dynamic Workbook) GetWorkbook(FileHandle handle)
    {
        if (handle.IsClosed)
            throw new InvalidOperationException($"File handle {handle.Id} is closed");
            
        if (!_handles.TryGetValue(handle.Id, out var excelWorkbook))
            throw new InvalidOperationException($"File handle {handle.Id} not found");
            
        return excelWorkbook;
    }
    
    internal static void Close(FileHandle handle, bool save)
    {
        if (!_handles.TryRemove(handle.Id, out var excelWorkbook))
            return;
            
        try
        {
            if (save) excelWorkbook.Workbook.Save();
            excelWorkbook.Workbook.Close(SaveChanges: false);
            excelWorkbook.Excel.Quit();
            
            // COM cleanup
            ComUtilities.Release(ref excelWorkbook.Workbook!);
            ComUtilities.Release(ref excelWorkbook.Excel!);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        finally
        {
            handle.IsClosed = true;
        }
    }
}
```

### 3. Core Commands Pattern

```csharp
// Example: RangeCommands
public class RangeCommands : IRangeCommands
{
    public async Task<OperationResult> GetValuesAsync(
        FileHandle handle, 
        string sheetName, 
        string rangeAddress)
    {
        return await Task.Run(() =>
        {
            var (excel, workbook) = FileHandleManager.GetWorkbook(handle);
            
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

### Phase 1: Add File Handle API (Parallel to Batch)

**Goal:** Introduce file handles without breaking existing code

1. Add `FileHandle` class to `ExcelMcp.Core/Models/`
2. Add `FileHandleManager` to `ExcelMcp.Core/Session/` (internal)
3. Add file commands to `IFileCommands`:
   - `Task<FileHandle> CreateAsync(string filePath)`
   - `Task<FileHandle> OpenAsync(string filePath)`
   - `Task SaveAsync(FileHandle handle)`
   - `Task CloseAsync(FileHandle handle)`
4. Update all `I*Commands` interfaces to accept `FileHandle` as first parameter
5. Keep old batch-based methods for backward compatibility (NO obsolete attribute - removes cleanly later)

**Example:**
```csharp
public interface IRangeCommands
{
    // NEW: File handle API
    Task<OperationResult> GetValuesAsync(FileHandle handle, string sheetName, string rangeAddress);
    
    // OLD: Batch API (keep during development, remove in Phase 4)
    Task<OperationResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress);
}
```

**Why no `[Obsolete]` attribute:**
- Obsolete warnings break build with `TreatWarningsAsErrors=true`
- Batch API remains fully functional during development
- Clean removal in Phase 4 (no gradual deprecation needed)

### Phase 2: Update MCP Server

**Goal:** Expose file handle API via MCP tools

1. Add new MCP tools:
   - `create_file(filePath)` → returns `{fileHandle: "abc123"}`
   - `open_file(filePath)` → returns `{fileHandle: "xyz789"}`
   - `save_file(handle)` → saves workbook
   - `close_file(handle, save?)` → closes workbook
2. Update all existing tools to accept `handle` parameter:
   - `excel_range(handle, action, ...)` - NEW
   - `excel_range(batchId, action, ...)` - OLD (deprecated)
3. Update tool descriptions to show file handle workflow
4. Create new prompt files showing simple open→work→close pattern

### Phase 3: Update CLI

**Goal:** Simplify CLI to match file handle model

1. Remove batch commands (`batch-begin`, `batch-commit`, etc.)
2. Add file commands:
   - `excelmcp file create <path>`
   - `excelmcp file open <path>`
   - `excelmcp file save <handle>`
   - `excelmcp file close <handle>`
3. Update all commands to require `--handle` flag:
   - `excelmcp range get-values --handle abc123 --sheet Sheet1 --range A1:D10`

**Alternative (Simpler):** CLI could auto-manage handles
```bash
# CLI opens, executes, saves, closes automatically
excelmcp range get-values --file report.xlsx --sheet Sheet1 --range A1:D10
```

### Phase 4: Remove Batch Infrastructure

**Goal:** Clean removal of obsolete batch code

1. Switch all consumers (MCP Server, CLI, tests) to file handle API (Phases 2-3)
2. Delete batch infrastructure in one clean commit:
   - `ExcelSession.cs`
   - `ExcelBatch.cs`
   - `IExcelBatch.cs`
   - Batch-based method overloads from all Commands interfaces
   - Batch MCP tool
   - Batch CLI commands
3. Remove batch-related documentation
4. Update CHANGELOG

**Why clean removal works:**
- No external consumers (Core is internal library)
- All internal consumers migrated in Phases 2-3
- No `[Obsolete]` warnings to clean up
- Single atomic change (easy to review)

## Benefits

### 1. Conceptual Simplicity
- ✅ Matches human mental model: "open file, work, save, close"
- ✅ No batch mode confusion
- ✅ No prompts/elicitations needed (LLMs already know this pattern)
- ✅ Clear lifecycle: create/open → work → save/close

### 2. Performance
- ✅ Excel stays open across operations (same as batch mode)
- ✅ No 2-17 second disposal between operations
- ✅ Fixes Issue #173 (sequential file lock errors)
- ✅ LLM controls when to close (vs automatic disposal)

### 3. Architecture
- ✅ Clear separation: Core manages handles, consumers use them
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

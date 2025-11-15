# Session API Redesign - Implementation Notes

Date: 2025-11-13
Branch: feature/session-api-redesign-spec

## Implementation Status

### ✅ Completed Foundation (Ready for Use)

1. **SessionManager in Core** (`src/ExcelMcp.Core/SessionManager.cs`)
   - Maps `sessionId` → `IExcelBatch` for both MCP Server and CLI
   - Methods: `CreateSessionAsync`, `GetSession`, `SaveSessionAsync`, `CloseSessionAsync`
   
2. **ExcelToolsBase Enhanced** (`src/ExcelMcp.McpServer/Tools/ExcelToolsBase.cs`)
   - Static `SessionManager` instance accessible via `GetSessionManager()`
   - New `WithSessionAsync` method for session-based operations
   - Legacy `WithBatchAsync` marked `[Obsolete]` but still functional

3. **excel_file Tool Updated** (`src/ExcelMcp.McpServer/Tools/ExcelFileTool.cs`)
   - ✅ `Open` action - creates session, returns `sessionId`
   - ✅ `Save` action - persists changes for active session
   - ✅ `Close` action - closes session without saving
   - ✅ Tool description updated with session lifecycle guidance

4. **Build Status**
   - ✅ Solution builds successfully (0 errors, 0 warnings)
   - ✅ Core and ComInterop unchanged (`IExcelBatch`, `batchId` preserved)

### 🚧 In Progress - Tool Migration Pattern

**Pattern for Migrating Tools:**

```csharp
// OLD (batchId + excelPath):
public static async Task<string> ExcelTool(
    ToolAction action,
    string excelPath,      // ❌ Remove
    string? param = null,
    string? batchId = null) // ❌ Change to sessionId

private static async Task<string> MethodAsync(Commands commands, string excelPath, string? param, string? batchId)
{
    var result = await ExcelToolsBase.WithBatchAsync(batchId, excelPath, save: false, ...);
}

// NEW (sessionId only):
public static async Task<string> ExcelTool(
    ToolAction action,
    string sessionId,      // ✅ Required session ID
    string? param = null)

private static async Task<string> MethodAsync(Commands commands, string sessionId, string? param)
{
    var result = await ExcelToolsBase.WithSessionAsync(sessionId, async (batch) => ...);
}
```

**Steps Per Tool:**
1. Update tool method signature: Remove `excelPath`, change `batchId?` → `sessionId` (required)
2. Update tool description: Remove batch guidance, add session requirement
3. Update all private method signatures: Replace `excelPath` and `batchId` with `sessionId`
4. Replace `WithBatchAsync` → `WithSessionAsync` in all private methods
5. Remove `save` parameter (sessions handle save via `excel_file` actions)

### 📋 Remaining Work

**Tools Needing Migration** (11 tools):
- ❌ `ExcelPowerQueryTool.cs` - 11 methods to migrate
- ❌ `ExcelWorksheetTool.cs`
- ❌ `ExcelRangeTool.cs`
- ❌ `ExcelTableTool.cs`
- ❌ `ExcelVbaTool.cs`
- ❌ `ExcelConnectionTool.cs`
- ❌ `ExcelDataModelTool.cs`
- ❌ `ExcelPivotTableTool.cs`
- ❌ `ExcelNamedRangeTool.cs`
- ❌ `ExcelQueryTableTool.cs`
- ❌ `BatchSessionTool.cs` - DELETE (replaced by excel_file Open/Save/Close)

**CLI Migration:**
- Add `session-open`, `session-save`, `session-close` commands
- Update all feature commands to accept `--session-id` instead of `--batch-id`
- Map CLI `--session-id` → Core `batchId` internally

**Testing:**
- Update MCP Server tests to use `excel_file` open/save/close workflow
- Add session lifecycle tests
- Verify backwards compatibility during transition

**Documentation:**
- Update all tool descriptions
- Update README with session workflow examples
- Update MCP prompts to explain session lifecycle
- Delete batch-related documentation

## Remaining TODOs

### 1. Core Layer - Keep `IExcelBatch` and `batchId` (NO CHANGES)
- **Core layer will NOT be changed** - keep all existing `IExcelBatch` interfaces and `batchId` parameters
- Core layer maintains its internal batch-based implementation
- Only MCP Server and CLI will expose the session abstraction to users

### 2. Introduce `SessionManager` and wire `excel_file` open/save/close (Core + MCP Server)
- Add a `SessionManager` in `src/ExcelMcp.Core` responsible for session lifecycle (shared by MCP Server and CLI):
  - Internal `ConcurrentDictionary<string, IExcelBatch>` `_activeSessions` (Core still uses IExcelBatch)
  - `CreateSession(string filePath)` → opens via `ExcelSession.BeginBatchAsync` and returns `sessionId`
  - `GetSession(string sessionId)` → returns `IExcelBatch?`
  - `SaveSession(string sessionId)` → calls `Save` on the batch
  - `CloseSession(string sessionId)` → calls `DisposeAsync`, removes entry
- Extend `FileAction` in `ExcelFileTool` with `Open`, `Save`, and `Close`
- Update `ExcelFile(...)` tool method to:
  - `open` → call `SessionManager.CreateSession(filePath)` and serialize a JSON response including `sessionId`, `filePath`, and `suggestedNextActions`
  - `save` → call `SessionManager.SaveSession(sessionId)` and return success/failure JSON
  - `close` → call `SessionManager.CloseSession(sessionId)` and return success/failure JSON (never saves)

### 3. Convert MCP tools to `sessionId` parameter (map to Core's `batchId`)
- For each MCP tool in `src/ExcelMcp.McpServer/Tools/**` (connections, datamodel, file, namedrange, pivottable, powerquery, querytable, range, table, vba, worksheet, etc.):
  - Remove `excelPath` parameter except for `excel_file` `open`/`create-empty` actions
  - Remove optional `batchId` parameter
  - Add `[Required] string sessionId` parameter (exposed to LLMs as "sessionId")
  - Inside tool methods, map `sessionId` → `batchId` when calling Core commands:
    ```csharp
    // MCP tool accepts sessionId from LLM
    public static async Task<string> ExcelPowerQuery(
        PowerQueryAction action,
        string? sessionId = null,  // ✅ Exposed to LLMs as sessionId
        ...)
    {
        // Get batch from session manager using sessionId
        var batch = SessionManager.GetSession(sessionId);
        if (batch == null) {
            return JsonSerializer.Serialize(new OperationResult {
                Success = false,
                ErrorMessage = $"Session '{sessionId}' not found",
                IsError = true
            });
        }
        
        // Core still uses IExcelBatch with its internal batchId
        return await batch.ExecuteAsync((ctx, ct) => {
            return await commands.ImportAsync(batch, queryName, mCodeFile);
        });
    }
    ```
- Ensure business errors are returned as JSON and `McpException` is only used for protocol/validation errors

### 4. Remove batch infrastructure (BatchSessionTool, WithBatchAsync, CLI BatchCommands)
- Delete `src/ExcelMcp.McpServer/Tools/BatchSessionTool.cs` and remove any references to it.
- Delete `src/ExcelMcp.CLI/Commands/BatchCommands.cs` and remove corresponding wiring and help text from `src/ExcelMcp.CLI/Program.cs`.
- Remove `ExcelToolsBase.WithBatchAsync` and any `if (batchId != null) { ... } else { ... }` dual-path logic.
- Rename any remaining `_activeBatches` fields to `_activeSessions` (or remove entirely if replaced by `SessionManager`).

### 5. Update CLI to use `--session-id` parameter (map to Core's `batchId`)
- Add or extend CLI file commands to mirror the MCP `excel_file` actions (`open`, `save`, `close`) and print/accept `sessionId` for scripting
- Update all feature CLI commands to accept `--session-id` parameter (exposed to users as "session-id")
- Inside CLI commands, map `--session-id` → `batchId` when calling Core:
  ```csharp
  // CLI accepts --session-id from user
  public int Import(string[] args)
  {
      var sessionId = args.GetOption("--session-id");  // ✅ CLI uses sessionId
      
      // Map sessionId → batchId for Core layer
      var task = Task.Run(async () => {
          await using var batch = await ExcelSession.BeginBatchAsync(
              filePath, batchId: sessionId);  // ✅ Core still uses batchId
          return await _commands.ImportAsync(batch, queryName, mCodeFile);
      });
  }
  ```
- Ensure CLI no longer creates temporary batches/sessions implicitly; instead, users/LLMs call `open` first, then operate, then `close`

### 6. Update tests for session lifecycle and adjust existing coverage
- ComInterop/Core tests:
  - **Keep all existing `IExcelBatch`/`BeginBatchAsync` references** - Core layer unchanged
  - Keep behavior and patterns identical (same Excel COM semantics, timeouts, etc.)
- MCP tests:
  - Rewrite flows that currently use `excel_batch` to use `excel_file` `open`/`save`/`close` and `sessionId`
  - Add tests for:
    - `excel_file(open)` returning a valid `sessionId`
    - `excel_file(save)` succeeds for a valid session
    - `excel_file(close)` closes without saving and returns success
    - Operations with missing/invalid `sessionId` return structured JSON errors
    - Read-only workflows: open → read → close (no save)
    - Discard workflows: open → modify → close (no save)

### 7. Update docs and prompts for session-only workflows
- Delete batch-related prompt artifacts:
  - `src/ExcelMcp.McpServer/Prompts/Content/excel_batch.md` (or equivalent).
  - `ExcelBatchModePrompts` usage and any batch-mode guidance.
- Update `excel_file` documentation and tool descriptions to:
  - Emphasize the required open/save/close lifecycle.
  - Make clear that `close` never saves and `save` is an explicit action.
- Update all tool `[Description]` attributes to:
  - Describe `sessionId` as a required parameter.
  - Remove references to `excelPath` for non-file tools.
- Update `README.md`, `examples/*.ps1`, `examples/*.sh`, and any relevant docs to show:
  - Default workflow: `open → operations(sessionId) → save → close`.
  - Read-only workflow: `open → read → close`.
  - Discard workflow: `open → modify → close` (no save).

---

This file captures the current implementation state of the Session API redesign and the remaining work needed to fully implement the "sessions only" Open/Save/Close model described in `specs/SESSION-API-REDESIGN-SPEC.md`.  
Use it as a checklist while completing the refactor across ComInterop, Core, MCP Server, CLI, tests, and documentation.

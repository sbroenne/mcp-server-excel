# Session API Redesign - Implementation Notes

Date: 2025-11-13
Branch: feature/session-api-redesign-spec

## Current State

### ComInterop (src/ExcelMcp.ComInterop)
- **NEEDS REVERT**: `IExcelBatch` was renamed to `IExcelSession` - needs to be reverted back to `IExcelBatch`
- **NEEDS REVERT**: `ExcelBatch` was renamed to `ExcelSessionInstance` - needs to be reverted back to `ExcelBatch`
- **NEEDS REVERT**: `ExcelSession.BeginBatchAsync` was renamed to `ExcelSession.OpenSessionAsync` - needs to be reverted back to `BeginBatchAsync`
- **Core layer should remain unchanged** - keep all existing `IExcelBatch` and `batchId` terminology

### Core (src/ExcelMcp.Core)
- **NEEDS REVERT**: Some command signatures were changed from `IExcelBatch` to `IExcelSession` in `Commands/DataModel/DataModelCommands.Read.cs` - needs to be reverted
- **All Core commands should remain unchanged** - keep all existing `IExcelBatch` and `batchId` parameters

### Strategy
- **Core and ComInterop layers**: Keep existing `IExcelBatch`, `BeginBatchAsync`, and `batchId` - no renames
- **MCP Server and CLI layers**: Expose `sessionId` parameter to users/LLMs, but internally map to Core's `batchId`
- This minimizes breaking changes while improving the user-facing API

## Remaining TODOs

### 1. Core Layer - Keep `IExcelBatch` and `batchId` (NO CHANGES)
- **Core layer will NOT be changed** - keep all existing `IExcelBatch` interfaces and `batchId` parameters
- Core layer maintains its internal batch-based implementation
- Only MCP Server and CLI will expose the session abstraction to users

### 2. Introduce `SessionManager` and wire `excel_file` open/save/close (MCP Server)
- Add a `SessionManager` in `src/ExcelMcp.McpServer` responsible for session lifecycle:
  - Internal `ConcurrentDictionary<string, IExcelBatch>` `_activeSessions` (Core still uses IExcelBatch)
  - `CreateSession(string filePath)` → opens via `ExcelSession.BeginBatchAsync` and returns `sessionId`
  - `GetSession(string sessionId)` → returns `IExcelBatch?`
  - `SaveSession(string sessionId)` → calls `SaveAsync` on the batch
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

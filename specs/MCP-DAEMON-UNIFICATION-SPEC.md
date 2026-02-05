# MCP Server Daemon Unification Specification

## Overview

Unify the MCP Server with the existing CLI daemon architecture to provide persistent session management across both interfaces.

## Problem Statement

### Current Architecture

```
┌──────────────────────────┐     ┌──────────────────────────┐
│     MCP Server #1        │     │     MCP Server #2        │
│  ┌────────────────────┐  │     │  ┌────────────────────┐  │
│  │  SessionManager    │  │     │  │  SessionManager    │  │
│  │  (isolated)        │  │     │  │  (isolated)        │  │
│  └────────────────────┘  │     │  └────────────────────┘  │
│           │              │     │           │              │
│     Excel Process A      │     │     Excel Process B      │
└──────────────────────────┘     └──────────────────────────┘
           ↑                                ↑
           │                                │
      File: test.xlsx ─────────────── File: test.xlsx
                        ❌ CONFLICT!
```

**Issues:**
1. Each MCP server process has its own `SessionManager`
2. Each opens separate Excel processes
3. File locking conflicts when multiple MCP servers access the same file
4. Sessions lost when MCP server process restarts
5. No visibility into sessions (no tray UI)

### CLI Daemon Architecture (Already Working)

```
┌─────────────────┐    Named Pipe    ┌──────────────────────────────────┐
│  CLI Command 1  │ ──────────────── │                                  │
└─────────────────┘                  │       ExcelDaemon                │
                                     │                                  │
┌─────────────────┐    Named Pipe    │  • SessionManager (singleton)    │
│  CLI Command 2  │ ──────────────── │  • Tray Icon                     │
└─────────────────┘                  │  • 10-min idle timeout           │
                                     │  • Single instance mutex         │
                                     │                                  │
                                     └──────────────────────────────────┘
```

**Benefits:**
- Single `SessionManager` across all CLI invocations
- Sessions persist between commands
- Tray UI shows active sessions
- Automatic cleanup via idle timeout

## Proposed Architecture

```
                                    ┌──────────────────────────────────┐
┌─────────────────┐                 │                                  │
│  CLI Commands   │──Named Pipe────▶│                                  │
└─────────────────┘                 │       ExcelDaemon                │
                                    │       (Unified)                  │
┌─────────────────┐                 │                                  │
│  MCP Server #1  │──Named Pipe────▶│  • SessionManager (singleton)    │
└─────────────────┘                 │  • Tray Icon (all sessions)      │
                                    │  • 10-min idle timeout           │
┌─────────────────┐                 │  • Single instance mutex         │
│  MCP Server #2  │──Named Pipe────▶│  • Core Commands                 │
└─────────────────┘                 │                                  │
                                    └──────────────────────────────────┘
                                                   │
                                           Excel Processes
                                        (one per open file)
```

**Benefits:**
1. ✅ Single `SessionManager` for CLI and MCP
2. ✅ No file locking conflicts between MCP instances
3. ✅ Sessions survive MCP server restarts
4. ✅ Unified tray UI shows all sessions
5. ✅ MCP Server becomes thin wrapper (less code to maintain)
6. ✅ LLM tests can use multiple turns without race conditions

## Implementation Plan

### Phase 1: Extract Daemon Client Library

**Goal:** Create reusable client library that both CLI and MCP can use.

**New Project:** `ExcelMcp.Daemon.Client`

```csharp
namespace Sbroenne.ExcelMcp.Daemon.Client;

public class DaemonClient : IDisposable
{
    public static DaemonClient Connect(bool autoStartDaemon = true);
    public Task<string> SendCommandAsync(string toolName, string action, Dictionary<string, object> parameters);
    public bool IsDaemonRunning { get; }
}
```

**Files to create:**
- `src/ExcelMcp.Daemon.Client/DaemonClient.cs`
- `src/ExcelMcp.Daemon.Client/DaemonProtocol.cs` (shared message types)
- `src/ExcelMcp.Daemon.Client/DaemonLauncher.cs` (auto-start logic)

### Phase 2: Refactor CLI to Use Client Library

**Goal:** CLI uses `DaemonClient` instead of direct pipe operations.

**Changes:**
- Extract pipe communication from `ExcelDaemon.cs` into shared protocol
- CLI commands use `DaemonClient.SendCommandAsync()`
- Verify existing CLI tests still pass

### Phase 3: MCP Server Uses Daemon

**Goal:** MCP Server tools forward requests to daemon.

**Before:**
```csharp
public class ExcelFileTool
{
    private static readonly SessionManager _sessionManager = new();
    
    public static string Open(string path, bool showExcel)
    {
        var batch = _sessionManager.CreateSession(path, showExcel);
        // Complex session management...
    }
}
```

**After:**
```csharp
public class ExcelFileTool
{
    public static async Task<string> Open(string path, bool showExcel)
    {
        using var client = DaemonClient.Connect();
        return await client.SendCommandAsync("excel_file", "open", new { path, showExcel });
    }
}
```

### Phase 4: Enhanced Tray UI

**Goal:** Tray shows session source (CLI vs MCP).

**Changes:**
- Track session origin in `SessionManager`
- Show in tray tooltip: "2 MCP sessions, 1 CLI session"
- Context menu: "Close all MCP sessions", "Close all CLI sessions"

## Protocol Design

### Request Format (JSON over Named Pipe)

```json
{
  "id": "uuid-v4",
  "tool": "excel_file",
  "action": "open",
  "parameters": {
    "excelPath": "C:\\test.xlsx",
    "showExcel": false
  },
  "source": "mcp-server"
}
```

### Response Format

```json
{
  "id": "uuid-v4",
  "success": true,
  "result": {
    "sessionId": "abc123",
    "filePath": "C:\\test.xlsx"
  }
}
```

### Error Response

```json
{
  "id": "uuid-v4",
  "success": false,
  "error": {
    "message": "File not found",
    "code": "FILE_NOT_FOUND"
  }
}
```

## Migration Strategy

### Backward Compatibility

1. MCP Server can work in **two modes:**
   - **Daemon mode** (default): Forward to daemon
   - **Standalone mode** (fallback): Use embedded `SessionManager`

2. Mode detection:
   ```csharp
   if (Environment.GetEnvironmentVariable("EXCELMCP_STANDALONE") == "true")
       UseEmbeddedSessionManager();
   else
       UseDaemonClient();
   ```

3. Auto-start daemon if not running (transparent to user)

### Testing Strategy

1. **Unit tests:** Mock `DaemonClient`, test tool logic
2. **Integration tests:** Start daemon, run MCP tests
3. **LLM tests:** Multi-turn workflows (the original problem!)

## File Structure Changes

```
src/
├── ExcelMcp.Daemon.Client/           # NEW: Shared client library
│   ├── DaemonClient.cs
│   ├── DaemonProtocol.cs
│   └── DaemonLauncher.cs
├── ExcelMcp.CLI/
│   ├── Daemon/
│   │   └── ExcelDaemon.cs            # MODIFIED: Use shared protocol
│   └── Commands/                      # MODIFIED: Use DaemonClient
├── ExcelMcp.McpServer/
│   └── Tools/                         # MODIFIED: Use DaemonClient
└── ExcelMcp.Core/                     # UNCHANGED
```

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Daemon startup latency | First call slow | Pre-launch daemon on install, lazy connect |
| Daemon crashes | All sessions lost | Robust error handling, reconnect logic |
| Protocol versioning | Breaking changes | Version field in protocol, negotiation |
| Security | Named pipe access | Keep existing security (per-user pipe) |
| Debugging complexity | Two processes | Unified logging, trace correlation |

## Success Criteria

1. ✅ MCP Server can complete 5-turn workflow without file locking
2. ✅ CLI and MCP sessions visible in same tray UI
3. ✅ Session survives MCP server restart
4. ✅ No performance regression (< 50ms added latency)
5. ✅ All existing tests pass
6. ✅ Standalone mode works for Docker/special cases

## Timeline Estimate

- **Phase 1:** 2-3 days (extract client library)
- **Phase 2:** 1-2 days (refactor CLI)
- **Phase 3:** 3-4 days (MCP integration)
- **Phase 4:** 1 day (tray enhancements)
- **Testing:** 2-3 days

**Total:** ~10-12 days

## Open Questions

1. Should daemon auto-start when MCP server connects?
   - **Recommendation:** Yes, with configurable behavior

2. Should we support multiple daemons (per-workspace)?
   - **Recommendation:** No, single daemon is simpler

3. Should daemon log to file or stdout?
   - **Recommendation:** File logging with rotation

4. What happens if daemon exits while MCP is running?
   - **Recommendation:** MCP reconnects and retries

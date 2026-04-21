---
name: "error-transport-context"
description: "Stamp MCP and CLI error responses with originating command and session context for consistent cross-transport diagnostics. Use when adding diagnostic fields to ServiceResponse, enriching error envelopes in ExcelToolsBase or CliErrorOutput, or debugging error context lost across the service boundary. Triggers: error handling, diagnostics, error envelope, ServiceResponse, error context, MCP error, CLI error, debug failure, error fields."
---

## Context

Use this when improving diagnostics in the MCP/CLI/service seam without changing Core or COM behavior.

## Patterns

1. Put shared diagnostics on `ServiceResponse`, not in host-specific wrappers.
2. Stamp failed responses with the originating `command` and `sessionId` at the service boundary so both MCP and CLI inherit the same context.
3. Mirror new shared fields in `ExcelToolsBase` and `CliErrorOutput` while keeping existing compatibility fields like `error`, `errorMessage`, and `isError`.
4. Prefer additive fields over message rewrites so existing clients keep working.

**Expected error JSON shape:**
```json
{
  "success": false,
  "isError": true,
  "error": "Table not found",
  "errorMessage": "Table not found",
  "command": "table.list",
  "sessionId": "abc123"
}
```

## Key Files

- `src/ExcelMcp.Service/ExcelMcpService.cs` — enrich failures in `ProcessAsync()` via a request-context helper.
- `src/ExcelMcp.McpServer/Tools/ExcelToolsBase.cs` — include `command` and `sessionId` in serialized error JSON.
- `src/ExcelMcp.CLI/Infrastructure/CliErrorOutput.cs` — emit the same fields in CLI failure envelopes.

## Verification

After changes, run existing integration tests to confirm no error field regressions:
```bash
dotnet test tests/ExcelMcp.Core.Tests --filter "RunType!=OnDemand"
```

## Anti-Patterns

- Do not add MCP-only diagnostics that CLI cannot emit.
- Do not push transport diagnostics down into Core commands just to label routed actions.
- Do not remove existing error fields that current tests or clients already consume.

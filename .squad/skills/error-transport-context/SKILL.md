---
name: "error-transport-context"
description: "Preserve service request context across MCP and CLI error envelopes."
domain: "error-handling"
confidence: "high"
source: "observed"
---

## Context
Use this when improving diagnostics in the MCP/CLI/service seam without changing Core or COM behavior.

## Patterns
- Put shared diagnostics on `ServiceResponse`, not in host-specific wrappers.
- Stamp failed responses with the originating `command` and `sessionId` at the service boundary so both MCP and CLI inherit the same context.
- Mirror new shared fields in `ExcelToolsBase` and `CliErrorOutput` while keeping existing compatibility fields like `error`, `errorMessage`, and `isError`.
- Prefer additive fields over message rewrites so existing clients keep working.

## Examples
- `src/ExcelMcp.Service/ExcelMcpService.cs`: enrich failures in `ProcessAsync()` via a request-context helper.
- `src/ExcelMcp.McpServer/Tools/ExcelToolsBase.cs`: include `command` and `sessionId` in serialized error JSON.
- `src/ExcelMcp.CLI/Infrastructure/CliErrorOutput.cs`: emit the same fields in CLI failure envelopes.

## Anti-Patterns
- Do not add MCP-only diagnostics that CLI cannot emit.
- Do not push transport diagnostics down into Core commands just to label routed actions.
- Do not remove existing error fields that current tests or clients already consume.

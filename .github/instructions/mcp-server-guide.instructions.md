---
applyTo: "src/ExcelMcp.McpServer/**/*.cs,src/ExcelMcp.Generators.Mcp/**/*.cs"
excludeAgent: "code-review"
---

# MCP boundaries

- Register with the SDK's `WithToolsFromAssembly`; it owns schemas, dependency
  injection, binding, transport, cancellation, and unexpected protocol errors.
  Do not introduce replacement registration or client-specific schema rewriting.
- Tool methods return `Task<CallToolResult>`. Inject the host-owned Service
  bridge and pass the SDK cancellation token explicitly through generated routes.
  Core commands remain synchronous on the owning Excel thread.
- Use `ExecuteToolActionAsync` for result conversion and telemetry. Execution
  failures set the actual SDK `CallToolResult.IsError`, not just a JSON field.
  Preserve structured Service error context and matching JSON text/structured
  content. Unexpected exceptions and cancellation propagate to the SDK.
- Request filters validate supplied action parameters and session identity
  before SDK binding. Generate action applicability from Core contracts.
- Ordinary shutdown saves confirmed remaining sessions; an unconfirmed Mac
  handoff stays untouched for manual recovery. Explicit no-save close discards
  edits; cancellation is not undo and must not close unrelated workbooks.
- Stdio stdout is JSON-RPC only, including startup/bootstrap paths; diagnostics
  go to stderr.
- Keep destructive/read-only tool metadata accurate.

Descriptions, input naming, server instructions, supported MCP capabilities,
and guidance validation follow
[Agent guidance sources](mcp-llm-guidance.instructions.md).

Manual routing example and explanation: `docs/DEVELOPMENT.md`.

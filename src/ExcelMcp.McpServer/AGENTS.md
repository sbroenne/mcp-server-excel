# MCP boundaries

Follow [repository review checks](../../AGENTS.md#code-review-rules) and
[runtime boundaries](../AGENTS.md). Owning MCP generators share these rules.

- Register with the SDK's `WithToolsFromAssembly`; it owns schemas, dependency
  injection, binding, transport, cancellation, and unexpected protocol errors.
  Do not introduce replacement registration or client-specific schema rewriting.
- Compiler diagnostics enforce the supported tool signature. Inject the
  host-owned Service bridge and pass cancellation through generated routes.
  Core commands remain synchronous on the owning Excel thread.
- Use `ExecuteToolActionAsync` for result conversion and telemetry. Execution
  failures set the actual SDK `CallToolResult.IsError`, not just a JSON field.
  Preserve structured Service error context and matching JSON text/structured
  content. Unexpected exceptions and cancellation propagate to the SDK.
- Request filters validate supplied action parameters and session identity
  before SDK binding. Generate action applicability from Core contracts.
- Ordinary shutdown closes saved owned workbooks and hands unsaved ones back
  open and visible without saving or discarding; attached workbooks stay untouched.
  Preserve AutoSave on open/attach. Explicit no-save close discards only unsaved
  edits; cancellation is not undo and must not close unrelated workbooks.
- Stdio stdout is JSON-RPC only, including startup/bootstrap paths; diagnostics
  go to stderr.
- Keep destructive/read-only tool metadata accurate.

Descriptions, input naming, server instructions, supported MCP capabilities,
and guidance validation follow
[Agent guidance sources](../../skills/AGENTS.md).

Manual routing example and explanation: [DEVELOPMENT](../../docs/DEVELOPMENT.md).

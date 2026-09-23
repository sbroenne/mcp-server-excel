---
applyTo: "src/ExcelMcp.McpServer/**/*.cs"
excludeAgent: "code-review"
---

# MCP boundaries

- Tool methods are static and synchronous. Use generated `ServiceRegistry`
  routes; hand-written branches are for special metadata or atomic no-session
  behavior, not duplicate dispatch.
- Scope the SDK cancellation token with `ExcelToolsBase.PushCancellationToken`;
  use `ExecuteToolAction` for telemetry/error handling and shared `JsonOptions`.
- Execution failures return structured JSON with `success: false` and
  `isError: true`. Invalid input/unknown actions may use established argument
  or protocol exceptions. Preserve Service error context rather than throwing
  a second generic exception.
- Stdio stdout is JSON-RPC only, including startup/bootstrap paths; diagnostics
  go to stderr.
- Descriptions add server-specific constraints and tool-selection hints, not
  types/enums already in the schema. No emojis in generated guidance/XML docs.
  Keep destructive/read-only metadata accurate.

Manual routing example and explanation: `docs/DEVELOPMENT.md`.

# Runtime boundaries

Follow [repository rules and review checks](../AGENTS.md).
These rules also apply to owning generators, not just hand-written runtime code.

- MCP calls its own `ExcelController`; CLI uses its named-pipe controller.
  Both route Excel work to private workers hosting `ExcelMcpService`, one
  execution owner per Excel application. MCP must not depend on the CLI daemon.
  Do not add Excel behavior to adapters or synchronous Excel work to controllers.
- CLI work that must continue after a command invocation exits belongs in the
  existing controller/worker lifetime, not a client-side background task. A
  background task does not survive its process exiting. Do not add client-side
  shutdown waits to compensate for the wrong owner. MCP owns corresponding
  work through its own controller and workers. See [process lifetimes](../docs/ARCHITECTURE.md#process-lifetimes)
  and [telemetry ownership](../docs/ADR-007-LOCAL-ACCESS-AND-TELEMETRY.md).
- Core commands are synchronous. Validate .NET inputs before `IExcelBatch.Execute`;
  COM work belongs on its STA callback. Do not wrap it in a catch returning a
  second result: batch/Service own failure transport and diagnostic context.
- ComInterop owns thread, session, and shutdown lifetime. Follow
  [COM safety](ExcelMcp.ComInterop/AGENTS.md) for Core and ComInterop changes.
  Optional-property probes must distinguish documented unavailable members from
  failures; an error, timeout, or cancellation must not become empty data.
  Follow [MCP boundaries](ExcelMcp.McpServer/AGENTS.md) for MCP and its generator.
- Public interface parameters use camelCase; generated MCP names use snake_case.
  Use naming attributes only for exceptions. Unknown action/enum values must
  fail rather than select a default.
- Never return credentials or unsanitized connection strings, even if an
  existing code path does so.

Examples and contributor conventions: [CONTRIBUTING](../docs/CONTRIBUTING.md).

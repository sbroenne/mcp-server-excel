# Runtime boundaries

- MCP calls `ExcelMcpService` in-process; CLI uses the named-pipe daemon.
  Do not add a daemon hop to MCP or Excel behavior to either adapter.
- CLI work that must continue after a command invocation exits belongs in the
  existing daemon, not a client-side background task. A background task does
  not survive its process exiting. Do not add a helper process or client-side
  shutdown waits to compensate for the wrong owner. MCP owns corresponding
  work in its own server process. See [process lifetimes](../../ARCHITECTURE.md#process-lifetimes)
  and [telemetry ownership](../../ADR-007-LOCAL-ACCESS-AND-TELEMETRY.md).
- Core commands are synchronous. Validate .NET inputs before `IExcelBatch.Execute`;
  COM work belongs on its STA callback. Do not wrap it in a catch returning a
  second result: batch/Service own failure transport and diagnostic context.
- ComInterop owns thread, session, and shutdown lifetime. Follow
  [COM safety](excel-com-interop.md) for COM changes.
- Public interface parameters use camelCase; generated MCP names use snake_case.
  Use naming attributes only for exceptions. Unknown action/enum values must
  fail rather than select a default.
- Never return credentials or unsanitized connection strings, even if an
  existing code path does so.

Examples and contributor conventions: [CONTRIBUTING](../../CONTRIBUTING.md).

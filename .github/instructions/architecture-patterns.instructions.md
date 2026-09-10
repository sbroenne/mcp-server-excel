---
applyTo: "src/**/*.cs"
excludeAgent: "code-review"
---

# Runtime boundaries

- MCP calls `ExcelMcpService` in-process; CLI uses the named-pipe daemon.
  Do not add a daemon hop to MCP or Excel behavior to either adapter.
- Core commands are synchronous. Validate .NET inputs before `IExcelBatch.Execute`;
  COM work belongs on its STA callback. Do not wrap it in a catch returning a
  second result: batch/Service own failure transport and diagnostic context.
- ComInterop owns thread, session, and shutdown lifetime. Follow
  `excel-com-interop.instructions.md` for COM changes.
- Public interface parameters use camelCase; generated MCP names use snake_case.
  Use naming attributes only for exceptions. Unknown action/enum values must
  fail rather than select a default.
- Never return credentials or unsanitized connection strings, even if an
  existing code path does so.

Examples and contributor conventions: `docs/CONTRIBUTING.md`.

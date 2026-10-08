# Runtime boundaries

- MCP calls `ExcelMcpService` in-process; CLI uses the named-pipe daemon.
  Do not add a daemon hop to MCP or Excel behavior to either adapter.
- Core commands are synchronous. Validate .NET inputs before `IExcelBatch.Execute`;
  COM work belongs on its STA callback. Do not wrap it in a catch returning a
  second result: batch/Service own failure transport and diagnostic context.
- ComInterop owns Windows thread, session, and shutdown lifetime. Mac session
  ownership lives in Service and must not terminate shared Excel. Follow
  [COM safety](excel-com-interop.md) and [Mac capability gates](../../../specs/MACOS-SUPPORT.md).
- Public interface parameters use camelCase; generated MCP names use snake_case.
  Use naming attributes only for exceptions. Unknown action/enum values must
  fail rather than select a default.
- Never return credentials or unsanitized connection strings, even if an
  existing code path does so.

Examples and contributor conventions: [CONTRIBUTING](../../CONTRIBUTING.md).

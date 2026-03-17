# Shiherlis — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Core Dev
- **Joined:** 2026-03-15T10:42:22.618Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: Bug report claim of a 13-column `set-values` limit does not match current Core behavior for rectangular payloads; live MCP repro and existing integration test both succeed beyond 13 columns.
- 2026-03-16: The current `SetValues` failure mode that matches the reported `ArgumentOutOfRangeException` is jagged input. `RangeCommands.SetValues` takes column count from row 0 and indexes later rows without validating rectangularity.
- 2026-03-16: Fresh-sheet writes to non-A1 ranges work in current MCP/Core when the payload shape matches the target range. Do not add sheet activation/select logic without a failing red integration test through `SheetCommands.Create` plus `RangeCommands.SetValues`.
- 2026-03-16: The safest Bug 1 fix is to validate each payload row against the resolved Excel range column count before building the 1-based COM array; that preserves valid wide writes and turns jagged input into a clear user-facing `ArgumentException`.
- 2026-03-16: `SetFormulas` had the same later-row indexing assumption as `SetValues`; mirroring the same width validation at that boundary closes the parallel defect without changing the broader range-write contract.

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
- 2026-04-08: Screenshot reliability for issues #563 and #583 improved when the target range is brought into view before `CopyPicture` and the copy uses `xlPicture`. Moving the temporary export chart to the range's workbook coordinates was not stable enough to keep; offscreen capture still needed clipboard/paste retry hardening because `chart.Paste()` can throw `0x800A03EC` if the export target is not in the visible viewport.
- 2026-04-08: The stable Core fix for screenshot regressions was to enforce visibility before `CopyPicture` and again before `chart.Paste()`: restore minimized Excel, activate the worksheet, `Goto` the range's top-left cell, scroll it into view, select the range, then activate/select the temp chart before pasting. Keeping the chart itself at `(0,0)` while targeting visibility through window scroll/selection proved more reliable than moving the chart object to offscreen range coordinates.

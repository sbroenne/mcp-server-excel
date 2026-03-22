# Cheritto — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Platform Dev
- **Joined:** 2026-03-15T10:42:22.620Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: Bug 3 is already supported under `range.set-number-format`; the actual gap is that the formatting surface is split between `range` (display formats) and `range_format` (visual styling), and top-level feature tables currently hide that split.
- 2026-03-16: Bug 5 is already supported under `range_format.auto-fit-columns` / `auto-fit-rows`; the main weakness is discoverability, not backend capability.
- 2026-03-16: Bug 4 is a real surface gap. A `List<...>` batch request is consistent with existing patterns such as `IRangeEditCommands.Sort(... List<SortColumn> ...)`, so the likely risk is in API review and docs parity, not generator feasibility.
- 2026-03-21: Stability bug (v1.8.32) isolates to CLI/daemon/service wrapper layer. Pure ComInterop session tests pass (Nate's serial workflow coverage), but CLI-through-daemon pathway lacks regression tests for serial PQ refresh → timeout → session recovery → new file open workflows. Test gap: no CLI integration tests exist for multi-step PQ refresh sequences that mirror the bug report's excelcli serial command pattern.
- 2026-03-21: **PIVOT DECISION REACHED.** Serial workflow control tests all GREEN confirms ComInterop/session timeout layer is correct. Any remaining bug lives in CLI daemon state, service routing, or workbook patterns. New mission: trace CLI/service/session wrapper seam for bug survival — identify where command routing, session wiring, or parameter propagation can fail without showing up in synthetic Core-layer regression tests.
- 2026-03-21: **Investigation Complete — Test Gap Documented.** Traced CLI-daemon-service seam and confirmed zero existing integration tests for serial Power Query workflows via CLI subprocess pattern. ComInterop control tests PASSED (7/7) — confirms pure batch/session layer stable. Created test pattern recommendations for `PowerQueryWorkflowTests.cs` with `CliProcessHelper` + `ServiceFixture` + workbook fixture for future regression testing. Identified specific wrapper-layer suspects: service RPC cancellation propagation, daemon session cleanup lazy evaluation, CLI process spawn → daemon connection continuity, and idle session monitor interference. **Recommendation: Keep test infrastructure ready, use recommended test pattern if future bug reports provide user workbook or specific workflows.**

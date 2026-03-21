# Nate — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Tester
- **Joined:** 2026-03-15T10:42:22.623Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: Bug triage has to start by checking the live tool surface; this report called out missing number formatting and auto-fit support even though `range.set-number-format` and `range_format.auto-fit-columns` already exist.
- 2026-03-16: Existing regression coverage can be directionally right but still too narrow; the wide-range write test only proves a 16-column single-row case, so multi-row threshold failures still need explicit regression tests.
- 2026-03-16: Workflow defects at the worksheet-create -> range-write boundary need MCP end-to-end coverage, because helper-created sheets can hide state bugs that only show up through the real tool sequence.
- 2026-03-16: Bug 1 now has a first red Core regression: a 14-column write with a shorter second row fails with `ArgumentOutOfRangeException`, confirming `RangeCommands.SetValues` does not validate rectangular row widths before indexing later rows.
- 2026-03-17: Bug 2 did not reproduce in Core integration when the workflow used `SheetCommands.Create` followed immediately by `RangeCommands.SetValues` to `A3:G10`; the payload round-tripped correctly and `A1` stayed empty, so any remaining defect signal is likely above the current Core command boundary.
- 2026-03-21: **Stability Investigation Completed.** Mapped 45+ existing tests across session, timeout, and lifecycle coverage. Identified serial Power Query workflow gap: create → refresh → modify → reopen sequence not tested as integrated flow. Proposed regression-test-first plan (baseline behavior → edge cases → recovery paths). Confirmed explicit timeout discipline for all stability tests.
- 2026-03-21: **Serial Workflow Regression Tests Added (7 tests, all GREEN).** Created `ExcelBatchSerialWorkflowTests` (3 tests) and `SessionManagerSerialWorkflowTests` (4 tests) modeling the v1.8.32 bug report's serial timeout recovery patterns. All tests PASS, confirming current implementation handles serial workflow recovery correctly: later operations fail fast (< 1s), dispose completes quickly (< 30s), Excel processes get killed, reopening same file works immediately, multiple sessions isolated from each other's timeouts. These are CONTROL tests, not RED tests — they validate correct existing behavior as regression baselines. Key finding: the bug report's workflow patterns are already handled correctly by the existing timeout infrastructure from PRs #525-#545.
- 2026-03-21: **PIVOT DECISION REACHED.** Serial workflow tests all GREEN confirms ComInterop/session timeout layer is healthy. Any remaining bug lives above this layer. New mission: find first true RED regression in CLI daemon state, service routing, or workbook-specific patterns. Investigation escalates to daemon-specific workflows (CLI parameter propagation, named-pipe state continuity) and workbook patterns (PQ data source failures, connectivity state).

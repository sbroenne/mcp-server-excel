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

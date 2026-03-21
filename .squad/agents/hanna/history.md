# Hanna — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** COM Interop Expert
- **Joined:** 2026-03-16T00:00:00Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: `Range.Value` documentation explicitly supports assigning a 2-dimensional array to a range in one operation, so validating payload shape before marshaling to COM is the correct boundary for `SetValues`.
- 2026-03-16: `Range.Formula2` documentation likewise allows a one- or two-dimensional Visual Basic array of matching dimensions, so mirroring the rectangular-width guard in `SetFormulas` is API-consistent rather than scope creep.
- 2026-03-16: `Range.Columns.Count` is a sound width source for the current contiguous-address write path, but Excel documents that it only reflects the first area of a multi-area range; named ranges that resolve to multi-area references should stay out of this Bug 1 fix and be handled only if a separate repro appears.
- 2026-03-21: **Stability Investigation Completed.** Reviewed COM API sequences in session/batch code under timeout. Findings: cleanup patterns are sound (try-finally, exponential backoff on Quit). Observable signals identified: QueryTable.Refresh behavior under timeout, Quit success/failure timing, ComUtilities.Release correctness, PID tracking. COM cleanup patterns support regression-first discipline. No code changes needed.

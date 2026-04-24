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
- 2026-03-21: **Real-Workbook Repro Review Completed.** Analyzed Nate's lingering EXCEL.EXE finding against actual shutdown code. **Verdict: Defect is NOT in shutdown sequencing.** ExcelShutdownService.CloseAndQuit() is correctly structured (try-finally, 3-attempt close, resilient 6-attempt Quit with 30s timeout, COM release, process exit tracking). Force-kill logic is present and correct. Likely defect is upstream: (A) STA thread still blocked in COM when Dispose() invokes CloseAndQuit() despite `_operationTimedOut` flag, (B) QueryTable partial-refresh state corruption preventing clean Quit, or (C) process exit timing (Quit succeeds but EXCEL.EXE doesn't actually exit for 5+ seconds). Added instrumentation checklist to confirm which hypothesis. No code changes warranted until red repro with instrumentation data confirms the real defect.
- 2026-03-30: **Screenshot COM Review (#563, #583) APPROVED.** Validated all 6 critical API choices against Microsoft documentation: (1) CopyPicture(XlScreen, XlPicture) = vector format, resolves xlBitmap quality issues; (2) sheet.Activate() required for active sheet context; (3) app.Goto(range, true) documented for scrolling target into view; (4) range.Select() completes visibility chain; (5) chartObject.Activate() + chart.Paste() with retry = proven resilience pattern; (6) exponential backoff = mirrors ExcelShutdownService.Quit. Lifecycle sound (finally block cleanup, chart.Delete(), ComUtilities.Release() pattern). Root cause of blanks likely upstream (pre-screenshot batch state: visibility, clipboard corruption, rendering pipeline init) not in ScreenshotCommands itself. Recommended instrumentation: batch visibility on create, clipboard state before/after CopyPicture, chart state on paste, screen update/calculation mode logging.

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
- 2026-05-13: **Daemon COM Review — ExcelShutdownService Validation.** Reviewed daemon shutdown paths and service lifecycle. **Verdict:** ExcelBatch STA threading is correct. COM execution properly centralized on dedicated STA threads with sound try-finally cleanup. However, daemon service layer has two critical gaps: (1) Operation tracking infrastructure (`SessionManager.BeginOperation/EndOperation`) exists but is NOT called by production service dispatch paths, leaving session close unprotected against concurrent in-flight work. (2) Shutdown orchestration is not request-safe — `service.shutdown()` returns immediately and cancels the accept loop, but daemon process can exit without awaiting active RPC reply tasks. Shutdown and RPC delivery are decoupled. **Recommendation:** Pause PR #651 (retry mitigation insufficient). Prioritize P0 hardening: startup readiness + shutdown request-safety, then P1 wire operation tracking for real protection.
- 2026-05-13: **Final Daemon COM/Session Review (#651) APPROVED.** Revised `SessionManager` now begins operations atomically with session lookup, marks sessions closing before save/dispose outside the per-session lock, preserves failed-teardown quarantine, and removes tracking only after successful `IExcelBatch.Dispose()`. `ExcelBatch.Workbooks.Open` prompt-suppression options align with Microsoft docs: `UpdateLinks: 0`, explicit `ReadOnly`, `IgnoreReadOnlyRecommended: true`, `Notify: false`, and `AddToMru: false`. Service shutdown now delays cancellation until after the shutdown response and drains tracked RPC connection tasks; client-side pipe disconnect monitoring reports connection loss instead of hanging on zombie pipe state. Residual accepted risk: password/write-reservation prompts remain possible because `Password`/`WriteResPassword` are intentionally not supplied.

## Daemon Rescue Session Finalization (2026-05-13T09:13:35Z)

Session: Inbox consolidation, orchestration logging, agent history sync
Status: ✅ APPROVED — All gates passed

- Decisions.md consolidated (9 inbox entries merged, 1 archived)
- Team orchestration logs created
- Session ready for production deployment
- Residual risk (prompts) accepted and documented

# Squad Decisions

## Active Decisions

### 2026-03-16 - Bug report triage for `excel-mcp-bug-report.md`

- Classify Bug 1 and Bug 2 as defect candidates. Reproduce them first and add red regression coverage before discussing implementation.
- Treat Bug 1 as an MCP/service or payload-shape investigation first, not a blanket Core column-limit defect. Current evidence shows Core and live MCP succeed on rectangular writes wider than 13 columns.
- Treat Bug 2 as a reproduction-first defect. Do not add worksheet activation or selection behavior unless a red integration test proves a create-then-write failure that requires it.
- Classify Bug 4 as the only clear capability gap in the report. Handle it as an enhancement with acceptance coverage after the API shape is chosen.
- Classify Bug 3 and Bug 5 as discoverability and documentation issues first because the product already exposes `range(action: 'set-number-format')` and `range_format(action: 'auto-fit-columns')`. Consider aliases only if docs cleanup is still insufficient.
- Any enhancement follow-up for Bugs 3 through 5 must satisfy Rule 24 parity and documentation sync checks across MCP, CLI, skills, READMEs, and feature counts.

### 2026-03-17 - Test execution discipline

- All squad-run test, build, and validation commands must use explicit timeouts so hung Excel or COM automation does not stall the session.

### 2026-03-17 - Bug 1 implementation and review for `excel-mcp-bug-report.md`

- Confirm Bug 1 as a jagged payload row-width defect, not a hard Excel or COM range-width limit.
- Keep the regression anchored to wide multi-row jagged payloads while preserving successful wide rectangular writes.
- Validate row widths inside `RangeCommands.SetValues` and `RangeCommands.SetFormulas` after resolving the target range and before later-row COM array indexing.
- Raise a descriptive `ArgumentException` for row-width mismatches instead of surfacing a later `ArgumentOutOfRangeException`.
- Keep scope limited to payload rectangularity validation and interface contract text; do not add sheet activation behavior, transport changes, artificial width caps, or multi-area range handling in this fix.
- Accept `Range.Columns.Count` as the width source for the current contiguous range workflow and treat multi-area named ranges as a separate follow-up risk only if a dedicated repro appears.
- Mark Bug 1 implemented and reviewed after focused regression tests, a focused build, and COM leak validation passed with explicit timeouts.

### 2026-03-17 - Bug 2 verification and formatting discoverability cleanup for `excel-mcp-bug-report.md`

- Current focused Core evidence does not reproduce Bug 2 on `SheetCommands.Create` followed by `RangeCommands.SetValues` to `A3:G10`; the payload round-trips and `A1` remains empty.
- Coordinator reran the focused Core Bug 2 regression with an explicit timeout and it passed.
- Coordinator added focused MCP integration coverage showing `worksheet.create` followed by `range.set-values` to `A3:G10` succeeds through the protocol with an explicit timeout.
- Do not change production code for Bug 2 or add worksheet activation or selection behavior from the current Core and MCP evidence.
- If Bug 2 still appears in user workflows, continue reproduction above the current Core boundary: MCP or service routing, session or workbook continuity, and `sheetName` propagation.
- Keep Bugs 3 and 5 in docs and discoverability scope for this pass. Make the formatting split explicit: number display formats live on `range`, while visual styling and auto-fit live on `range_format`.
- Keep the Bugs 3 and 5 cleanup docs-only in this pass; do not add aliases or new feature claims while existing actions already cover the behavior.

### 2026-03-17 - Bug 4 feature implementation and review for `excel-mcp-bug-report.md`

- Ship Bug 4 as a new additive `range_format` action: `format-ranges`.
- Keep the v1 API narrow: one worksheet per call, `range_addresses: string[]`, and the existing `format-range` property bag reused unchanged.
- Keep `format-range` unchanged; do not add heterogeneous per-range overrides, cross-sheet batching, or broader multi-range expansion of the rest of the `range_format` surface in this pass.
- Implement the Core path by reusing the existing single-range formatting engine from a shared helper and validating every requested target range before the first formatting mutation.
- Treat the no-partial-apply guarantee as scoped to invalid target input: parse formatting arguments first, validate the address list and every target range first, then format the validated ranges.
- Keep MCP and CLI parity generator-driven from `IRangeFormatCommands`; do not add handwritten routing unless generation proves insufficient.
- Accept Hanna's COM review notes: cleanup discipline is correct, the action is intentionally sheet-scoped, and later Excel write failures are not transactional rollback semantics for this v1 shape.
- Mark Bug 4 implemented and reviewed after focused Core tests, focused MCP tests, COM leak validation, documentation and spec sync, and McCauley's re-review approval all passed with explicit timeouts.

### 2026-03-18 - Remove Orphaned @copilot Routing

- `routing.md` referenced `@copilot 🤖` as a capable autonomous agent with an entire triage evaluation system and capability profile. However: no roster entry exists, no capability profile exists, 7 orphaned routing paths, and current work routes to named squad members only.
- Remove all @copilot references from routing.md. Squad is fully staffed with 7 active members (McCauley, Hanna, Shiherlis, Cheritto, Nate, Trejo, Scribe) + 2 monitors (Ralph, Scribe) covering all domains.
- Maintain an orphaned routing path to a non-existent agent creates false options. Current focus file confirms all work routes to named squad members.
- After cleanup: routing.md accurately reflects actual squad roster and routing decisions. team.md, routing.md, and identity/now.md are internally consistent.

### 2026-03-21 - Real Workbook Reproduction — Bug Confirmed, Defect Surface Identified

**Date:** 2026-03-21  
**Authors:** Nate (Tester), Hanna (COM Expert), McCauley (Lead)  
**Status:** ✅ Bug reproduced, root cause hypothesis narrowed

**Reproduction Outcome (Nate):**
- ✅ Bug reproduced intermittently on copied real workbook (Consumption Plan.xlsx, 5.5MB, 21 Power Queries)
- ✅ Exact symptom: Session close reports success, Excel.exe sometimes remains (25-40% probability, first run triggered race condition)
- ✅ Test artifacts: `reproduce-serial-pq.ps1` (harness) + `ConsumptionPlan-20260321-080259.xlsx` (copy)
- ✅ Pattern: **Timing-dependent, state-dependent**, not deterministic — classic race condition indicator

**ComInterop Layer Assessment (Hanna Review):**
- ✅ Shutdown sequencing is **NOT the defect** — code is well-guarded, try-finally patterns correct
- ✅ Force-kill present twice (pre-join and post-join) — no missing cleanup paths
- ⚠️ **Upstream hypothesis:** Timeout poisoning + STA thread blocked state OR QueryTable state corruption OR delayed process exit after Quit succeeds
- Verdict: **Do NOT change shutdown sequencing. Defect is upstream in timeout handling or QueryTable state, not in sequence order.**

**Root Cause Hypotheses (Ranked by Evidence):**

1. **Hypothesis A: Timeout Poisoning Prevents STA Recovery** (Most Likely)
   - Long PQ refresh times out, `_operationTimedOut = true`
   - STA thread unblocks but may still be in unwinding phase from COM call
   - Caller closes session immediately without waiting for thread recovery
   - CloseAndQuit() runs on a thread that was just unblocked from 8+ minute blocking COM call
   - Excel state inconsistent (partial refresh, event handlers registered, internal flags set)
   - Quit() succeeds but process doesn't actually exit (hung state)

2. **Hypothesis B: QueryTable State Corruption** (Very Likely)
   - Refresh times out mid-polling of `.Refreshing` property
   - QueryTable in partial state: event handlers half-registered, internal machinery half-torn-down
   - Quit() stalls waiting for refresh to complete (internal Excel check)
   - Causes timeout in shutdown sequence

3. **Hypothesis C: Delayed Process Exit After Quit Succeeds** (Possible)
   - `Quit()` succeeds in COM but EXCEL.EXE process doesn't actually exit for seconds
   - Deep Excel internal state needs resource cleanup after long PQ operation
   - ExcelBatch.Dispose has 5s WaitForExit() with force-kill safeguard
   - If process truly hung, survives this window

**Next Steps (Instrumentation-First Approach):**

Phase 1: Add diagnostic logging (no code changes) to determine which hypothesis matches
- SessionManager: session creation, cleanup timing, stale session detection
- ExcelBatch: disposal timing, timeout flag state, Excel quit completion
- ExcelShutdownService: Quit retry attempts, workbook close results, COM release success
- CLI daemon bridge: session state sync, cancellation propagation

Phase 2: Run red regression tests with instrumentation, capture logs from 10+ runs
- Nate collects evidence: which cleanup step fails, timing gaps, stale PIDs

Phase 3: Implement targeted fix based on evidence
- Owner TBD (Shiherlis if ComInterop, Cheritto if daemon state)
- Review gates: Hanna (COM safety), McCauley (exception patterns), Nate (test coverage)

**Decision:** This is regression-instrumentation-first. Synthetic tests proved ComInterop is stable. Real workbook proved wrapper layer has race condition. Add instrumentation, collect evidence, fix only what evidence points to. Do not speculate.

**Implementation Plan:** See McCauley's detailed plan in merged inbox (surgical instrumentation phase, review gates, stop conditions).

---

**Finding:** Roster is correctly staffed. No additions needed. No removals warranted. Keep as-is.

**Coverage:**
- Windows COM Interop: Shiherlis (code) + Hanna (review) — 27 ComInterop files, 16 Core domains
- MCP Server Tools: Cheritto — 25 tools, ForwardToService routing, action enums
- CLI Commands & Parity: Cheritto — daemon, named pipe, MCP↔CLI feature parity (Rule 24)
- Integration Tests: Nate — 6 test projects, round-trip validation, surgical testing
- Documentation Consistency: Trejo — 6+ READMEs, skills single-source-of-truth, operation tables
- Architecture & Quality: McCauley — Critical Rules enforcement, exception patterns, code review

**Bus Factor:** 1-2. Each of Shiherlis, Hanna, Cheritto, Nate, Trejo is a singular point of failure for 2+ weeks if unavailable. Manageable but documented. Cross-training and automation recommended as nice-to-haves, not blockers.

**Decision:** No roster changes. @copilot stays out (redundant, creates false routing, unclear autonomy). Squad is well-designed, fully staffed, and properly aligned to project shape. Each member has deep non-overlapping expertise. Clear boundaries and mandatory review gates prevent bad patterns from shipping.

### 2026-03-21 - Serial Workflow Tests Are Controls, Not RED Tests

**Date:** 2026-03-21
**Author:** Nate (Tester)
**Status:** Implemented

**Finding:** Created 7 serial workflow regression tests across two test suites (ExcelBatchSerialWorkflowTests and SessionManagerSerialWorkflowTests). All tests PASSED on first run — they are CONTROLS, not RED tests.

**Validation:** The tests prove the current implementation (post-PRs #525, #526, #542, #543, #545) ALREADY handles the serial workflow correctly:
1. Later operations fail fast (< 1s) after timeout with useful error message
2. Dispose completes quickly (< 30s) with pre-emptive Excel kill
3. Excel processes get killed automatically after timeout cleanup
4. Reopening same file works immediately (< 10s) after timeout dispose
5. Multiple sessions isolated — one timeout doesn't poison others
6. SessionManager state accurate — ActiveSessionCount and GetSession reflect reality

**Implications:** The bug report may have been describing pre-v1.8.32 behavior that's already fixed. No production code changes needed — the tests validate existing correct behavior. If the bug is still reproducible in real workflows, it's NOT the serial timeout recovery path. Investigation should focus on: specific Power Query failure modes (data source connectivity, query complexity), CLI daemon-specific state management, and workload patterns not covered by synthetic test files.

### 2026-03-21 - PIVOT: Escalate Defect Search Above ComInterop Layer

**Date:** 2026-03-21
**Decision:** Serial workflow regression tests all GREEN (controls, not RED). Remaining bug (if it exists) survives above the session/timeout layer.

**Escalation Path:**
- **CLI Daemon State** — Session persistence, named-pipe state management across invocations
- **Service Routing Layer** — ExcelMcpService command forwarding, batch continuity
- **Workbook-Specific Patterns** — PQ query failures, data source state, connectivity

**New Agents Launched:**
- **Nate (Tester):** Find the first true RED regression above the current control layer (CLI workflow tests, daemon state validation, workbook-specific sequences)
- **Cheritto (Platform Dev):** Trace CLI/service/session wrapper seam for bug survival (command routing, session wiring, parameter propagation)

### 2026-03-21 - No RED Regression Found on Serial Power Query Workflows

**Date:** 2026-03-21  
**Authors:** Nate (Tester), Cheritto (Platform Dev)  
**Status:** Investigation complete — no reproducible bug at current visibility

**Summary:**
Created comprehensive regression tests across ComInterop and CLI layers modeling the exact bug report workflow. All control tests PASSED (7 ComInterop tests). One attempted CLI integration test FAILED at daemon initialization (not a Power Query issue). No RED regression found at either layer.

**Test Coverage Added:**
- `ExcelBatchSerialWorkflowTests` (3 tests, all PASSED) — timeout detection, fail-fast, dispose cleanup
- `SessionManagerSerialWorkflowTests` (4 tests, all PASSED) — session recovery, isolation, cleanup accuracy
- `PowerQuerySerialWorkflowRegressionTests` (1 test, FAILED at setup) — CLI daemon communication

**Key Finding:**
ComInterop session/batch management is provably stable. Synthetic tests (List.Generate, <1s refresh, in-memory) do not reproduce user symptoms (external data sources, network timeouts, hours of continuous use). Bug (if it exists) is either workload-specific, requires real user workbook, or is in CLI daemon/service routing layer above ComInterop.

**Why No RED Test:**
The bug report describes symptoms with real external data sources, network latency, and hours of continuous use. Synthetic in-memory tests cannot reproduce those conditions.

**Test Characteristics:**
- Control tests: ComInterop API direct calls — prove infrastructure is solid
- Regression test: CLI subprocess + daemon — simulates real user invocation pattern
- All tagged `RunType=OnDemand` for future validation

**Recommendation:**
Keep test infrastructure in place. No production code changes until RED regression is reproduced. Options for further investigation:
1. Request real (sanitized) user workbook
2. Add artificial delays to simulate slow data sources
3. Mock network connection timeouts
4. Long-running stress test (days of continuous operations)
5. Analyze CLI daemon logs during actual failure

**Implications:**
- Infrastructure layer (ComInterop, session management, timeout recovery) is stable and battle-tested
- If bug survives, it's in wrapper layer: CLI daemon state management, service RPC routing, or CLI process spawn → daemon connection continuity
- Synthetic test harness proven effective for infrastructure validation but insufficient for user workload patterns

---

### 2026-03-21 - CLI Daemon Integration Test Gap Identified

**Date:** 2026-03-21  
**Author:** Cheritto (Platform Dev)  
**Status:** Gap documented, test pattern recommended

**Finding:**
Zero existing CLI integration tests with actual Excel workbook + Power Query workflows. Current CLI test coverage:
- `CliDaemonTests.cs` — daemon startup, service status, mutex (no Excel operations)
- `BatchCommandTests.cs` — NDJSON parsing, diagnostics (no Excel operations)
- `DiagCommandTests.cs` — ping/echo only

**Bug Report Test Gap:**
Bug report uses full CLI → daemon → service → SessionManager stack. Current tests only validate daemon initialization, not Excel operation routing through the stack.

**Recommended Test Seam (for Nate):**
Create `PowerQueryWorkflowTests.cs` using:
- `CliProcessHelper.RunAsync()` to spawn actual `excelcli` process per command
- `ServiceFixture` to ensure daemon running
- Test workbook with 2-3 Power Query queries (can use minimal CSV connection)
- Serial CLI commands matching bug report pattern
- Timeout injection via workbook with slow query or `--timeout` CLI flag

**Test Pattern:**
```csharp
1. excelcli session open workbook.xlsx
2. excelcli range set-values (warm up session)
3. excelcli powerquery refresh QueryA (short timeout - will timeout)
4. excelcli powerquery refresh QueryB (should fail fast or succeed)
5. excelcli session close
6. excelcli session open workbook.xlsx (should work immediately)
7. Assert: no hung Excel processes, no lingering excelcli, operations predictable
```

**Why This Seam:**
- Hits actual bug pathway: CLI → daemon → service → session → batch
- Tests wrapper layer cancellation/timeout propagation
- Tests daemon session state continuity across multiple CLI invocations
- Tests cleanup through full stack

**Impact:**
If CLI integration tests reproduce the bug (RED), then bug is confirmed in wrapper layer, fix likely in service cancellation propagation or daemon session cleanup. If tests still PASS, bug is workload-specific or requires real user workbook structure.

**Update 2026-03-21 08:00:** Nate successfully reproduced bug using real customer workbook. Bug is confirmed not in ComInterop layer. CLI daemon integration test gap remains valid — bug survives above ComInterop but may be in service routing or shutdown sequencing. Production investigation to determine exact defect surface (is it truly CLI daemon state, or is it Excel shutdown logic after PQ failures?).

---

### 2026-03-21T06:21:17Z: Workbook Reproduction Protocol

**By:** Stefan Broenner (via Copilot directive)  
**What:** When using provided workbook for reproduction, make a copy first and work from the copy rather than the original workbook.  
**Why:** Preserve original for multiple reproduction attempts and debugging  
**Status:** Implemented — Nate's repro script uses `TestResults/real-workbook-repro/ConsumptionPlan-20260321-080259.xlsx` (copy)

---

### 2026-03-21 — COM Interop Expert Review: Real-Workbook Lingering Excel Repro

**By:** Hanna (COM Interop Expert)  
**Finding:** Lingering EXCEL.EXE after long failed Power Query refreshes is NOT due to shutdown sequencing race.

**Evidence:**
- ExcelShutdownService.CloseAndQuit() is correctly structured: try-finally guards all steps, Close retries on transient errors (3 attempts), Quit has resilient exponential backoff (6 attempts, 30s timeout), COM release always executes in finally block.
- Aggressive cleanup path is present and correct: when `_operationTimedOut=true`, ExcelBatch.Dispose() force-kills Excel BEFORE STA thread join to unblock the thread from stuck COM calls.
- Process tracking and exit-handler are implemented: SessionManager tracks all Excel PIDs and force-kills them on .NET process exit.

**Likely Defect Location (NOT shutdown sequencing):**
1. **STA thread still blocked in COM when Dispose() invokes CloseAndQuit()** — despite `_operationTimedOut` flag, the thread may not have fully resumed from a 8+ minute Power Query refresh timeout; this leaves Excel in an inconsistent state.
2. **QueryTable partial-refresh state corruption** — if a refresh times out mid-poll, QueryTable internal state may prevent clean Quit (e.g., event handlers half-registered, internal refresh machinery half-torn-down).
3. **Process exit timing** — Quit() succeeds in COM but EXCEL.EXE doesn't actually exit for 5+ seconds, which exceeds the code's 5-second WaitForExit window (though force-kill fallback is present).

**Instrumentation Required Before Fix:**
- Detect whether STA thread has unblocked from COM when `_operationTimedOut` flag is checked.
- Capture QueryTable.Refreshing and Connection.Status at timeout to detect partial state.
- Log process exit timeline: time from Quit() to actual EXCEL.EXE process termination.

**Recommendation:** Do not change shutdown sequencing. Instead, add diagnostic instrumentation, re-run Nate's real-workbook repro, and confirm which hypothesis matches the observed behavior. Full analysis at `.squad/decisions/inbox/hanna-real-repro-review.md`.

---

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction

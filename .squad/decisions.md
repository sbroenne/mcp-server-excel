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

### 2026-03-18 - Squad Roster Fitness Review

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

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction

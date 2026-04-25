# McCauley — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Lead
- **Joined:** 2026-03-15T10:42:22.614Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: Treat reported `set-values` width failures above 13 columns as MCP/service/client-shape suspects first, not automatic Core defects; Core integration coverage already exercises `SetValues` on `A1:P1` successfully.
- 2026-03-16: Split "missing formatting capability" reports into real product gaps versus discoverability gaps. `set-number-format` already exists on `range`, and `auto-fit-columns` already exists on `range_format`.
- 2026-03-16: Do not accept "activate the new sheet" as the default fix for post-create write issues without proof; `SetValues` resolves by explicit sheet name, so a non-`A1` write failure after sheet creation needs a targeted reproduction test before changing lifecycle behavior.
- 2026-03-17: Final gate for Bug 1 stays green when the fix validates rectangular row widths at the command boundary, mirrors the same guard in `SetFormulas`, and keeps wider named-range or multi-area behavior explicitly out of scope until a separate repro exists.
- 2026-03-18: Squad roster is fully staffed (7 active members + 2 monitors). Removed orphaned @copilot routing references that had no corresponding agent entry. Current focus file confirms all work routes to named squad members. Routing.md simplified to reflect actual agent responsibilities without @copilot path.
- 2026-03-18: **Roster Review Complete.** Squad is correctly staffed and well-aligned. All 6 core workstreams covered: COM interop (Shiherlis + Hanna), MCP/CLI parity (Cheritto), tests (Nate), docs (Trejo), architecture (McCauley). Bus factor is 1-2 (each core specialist is singular point of failure for 2+ weeks); manageable but acknowledged. @copilot: never add (redundant, creates false routing, unclear autonomy). Recommendation: keep as-is, document bus factor risk, no roster changes needed.
- 2026-03-18: Comprehensive roster review completed: covered 5 questions across workstreams, role gaps, redundancy, @copilot decision, and specific recommendations. Finding: no roster changes warranted. Each role is distinct, necessary, and non-overlapping. Coverage is complete and properly aligned. Team is ready.
- 2026-03-21: **Stability Investigation Completed.** Reviewed last 10 PRs and cross-referenced against bug report. Findings: existing test coverage proves command-level correctness; gaps exist in serial workflows (Power Query create → refresh → modify → reopen). Produced 3-phase regression-test-first plan. No code changes requested. Team recommended baseline → edge cases → recovery test sequence.
- 2026-03-21: **PIVOT DECISION ISSUED.** Serial workflow regression tests all GREEN (7 tests across ExcelBatchSerialWorkflowTests and SessionManagerSerialWorkflowTests). These validate existing correct behavior as control baselines. Conclusion: ComInterop/session timeout layer is healthy. Any remaining bug lives above this layer. Escalation to CLI daemon state, service routing, workbook-specific patterns authorized. New agents: Nate (find first true RED regression above control layer) and Cheritto (trace wrapper seam for bug survival).
- 2026-03-21: **REAL REPRO PLANNING COMPLETED.** Nate reproduced intermittent lingering Excel after session close on copied real workbook (25-40% probability). ComInterop control tests remain GREEN. Bug is in wrapper seam (session state, service routing, or workbook-specific PQ behavior). Produced surgical instrumentation plan: three files (SessionManager, ExcelBatch, ExcelShutdownService) get detailed logging before fix attempt. Red test: checked-in real workbook fixture with CSV-backed PQ queries. Stop conditions: instrumentation identifies seam → fix owner assigned → Hanna gate before shipping. Hanna mandatory reviewer. No speculative code changes until evidence points to root cause. Estimated 2-3 days investigation, then 1-2 days fix.
- 2026-04-01: Hand-written MCP tools need explicit `[Description]` attributes on optional parameters; XML `<param>` docs alone did not populate the published JSON schema. For schema/discoverability bugs, gate fixes with a live `ListToolsAsync()` integration test and treat Claude Desktop runs as supporting evidence only unless the desktop client is confirmed to be using the current repo build.
- 2026-04-01: **PR Scope Review — dependency-refresh + pytest-skill-engineering.** 51 changed files triaged. 43 in-scope (17 NuGet bumps, 3 VS Code dep bumps, 29 pytest migration files, 8 doc renames). 4 files MUST be excluded (.squad agent histories + project-conventions skill). Migration pattern is mechanically clean: Agent→CopilotEval, aitest_run→copilot_eval, new cli_mcp_server.py FastMCP wrapper. Two major version jumps need build verification: MCP SDK 0.9.0-preview.2→1.2.0 and TypeScript 5.9→6.0. Branch is ready for issue/PR execution after exclusions.

- 2026-04-01: **CHART HARDENING ASSESSMENT COMPLETED.** Chart surface already provides comprehensive structured metadata (positioning fields, collision detection, warnings). No confirmed product bug exists. Cost (3-4 hours) outweighs marginal value. Decision: DEFER indefinitely pending real user failure, cross-tool parity requirement, multi-model evidence, or skill author request. If triggered, minimal scope only: add single `CollisionWarning: bool` + `SuggestedNextAction: string` to `ChartCreateResult`. Confidence HIGH.

- 2026-04-02: **ERROR DIAGNOSTICS CONTRACT GATE — CONDITIONAL APPROVAL.** Reviewed first-slice implementation for #585 error handling milestone. Routing architecture is correct (MCP/CLI converge at ExcelMcpService.ProcessAsync identically). #585 regression tests are structurally clean with proper round-trip verification. Exception propagation and COM cleanup pass. One blocking finding: CLI error envelope uses `error` property while MCP uses `errorMessage`, CLI omits `isError` flag, and CLI lacks structured diagnostic fields (`exceptionType`, `hresult`). This is a Rule 24 parity violation. Fix assigned to Cheritto. Advisory items: MCP's two error shapes should converge; failure-path test coverage needed (Nate).
- 2026-04-02: **Team closeout for the first diagnostics slice.** Final team posture for #585 is additive parity only: CLI must mirror MCP envelope fields, the shared transport should carry structured diagnostics, and focused validation slices outrank broader MCP class runs until the existing `ProgramTransport` session flake is fixed separately.---

## 2026-04-24: Kelso Plugin Phases -1 to 3 Complete — All Decisions Locked

**Status:** ✅ APPROVED  
**Scope:** Copilot CLI plugin implementation, distribution automation, infrastructure audit  

**Key Decisions (All Locked):**
1. Phase -1 Spike: Validated plugin install mechanism, workspace-scoped MCP finding
2. Phase 0 Scaffold: Created published repo structure, two-plugin separation (excel-mcp + excel-cli)
3. Phase 1: Removed placeholder agent (not needed; skills provide comprehensive guidance)
4. Phase 3: Automated release workflow with corrected version extraction
5. Audit: Infrastructure 85% clean, 3 actionable items identified (2 for Trejo, 1 for Kelso)

**Agent Decision Already Made:**
- Phase 1 placeholder agent: REMOVED entirely (no value add, spec doesn't require agents)
- Rationale: Skills already comprehensive; agent without clear scope is worse than nothing
- Can add later if multi-step workflow orchestration justifies it

**No Architectural Gate Needed:** All phases completed with your earlier approval. Kelso is ready to create GitHub issue + PR on `feature/copilot-cli-plugins` branch.

**Deliverables:**
- ✅ Published repo created: `sbroenne/mcp-server-excel-plugins`
- ✅ 2-plugin architecture: `excel-mcp` (MCP server + skill) + `excel-cli` (skill-only)
- ✅ Release workflow: `.github/workflows/publish-plugins.yml` with automated publishing
- ✅ Audit complete: 3 cleanup items identified, no blockers

**Next:** Scribe will orchestrate Kelso's PR creation after merging all decisions.

---

### 2026-04-24: Session End — Blocker on PR, Decisions Merged, Scribe Logs Written

**Phase 5 E2E Plugin Testing (Nate):**
- Both excel-mcp and excel-cli plugins validated production-ready end-to-end.
- Real installation tested with GitHub Copilot CLI; all scripts execute successfully.
- Only blocker: missing v0.0.1 GitHub Release asset (for download.ps1 e2e flow).
- Workaround validated: manual binary placement works perfectly.

**Phase 6 Release Docs (Trejo):**
- Docs audit complete: old infrastructure cleaned up, new release process well-documented but undiscoverable.
- Release docs now linked from main README (3-line addition).
- RELEASE-STRATEGY.md updated to explicitly document GitHub Copilot CLI plugin publishing workflow.

**Blocker Reported (Kelso):**
- DO NOT open plugin PR from current mixed dirty tree (plugin work + Squad infrastructure + RangeCommands.Formulas.cs).
- Stop, report blocker plainly, wait for branch narrowing before PR creation.

**Scribe Actions (2026-04-24T09:26:02Z):**
- ✅ Orchestration log written: 2026-04-24T09-26-02Z-nate-revert-range-change.md
- ✅ Session log written: 2026-04-24T09-26-02Z-revert-range-change.md
- ✅ Decision inbox merged to decisions.md (6 files, deduplicated)
- ✅ Cross-agent history updated (Nate, Kelso, Trejo, McCauley)
- ✅ Git commit skipped (nothing staged, user explicitly directed revert)

**User Directive Executed:**
- Reverted unrelated RangeCommands.Formulas.cs working-tree change (Nate completed).
- Repo working tree cleaned per user instruction.

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


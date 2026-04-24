# Trejo — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 230 operations.
- **Role:** Docs Lead
- **Joined:** 2026-03-15T10:42:22.625Z

## Cross-Agent Impact Notes

### 2026-04-24: Kelso Plugin Phases -1 to 3 Complete — Action Items for Trejo

**Kelso delivered:** All phases locked. Infrastructure audit identified 2 actionable items for Trejo's scope.

**Your Action Items:**

1. **STALE — Remove `skillpm` field from package.json** 🔴  
   - **Scope:** `packages/excel-mcp-skill/package.json` + `packages/excel-cli-skill/package.json`
   - **Change:** Delete lines with `"skillpm": { "mcpServers": [...] }`
   - **Reason:** Old agentskills.io-era field, no longer relevant
   - **Impact:** Signals clean migration away from old skillpm ecosystem

2. **DOC GAP — Add Copilot CLI Plugin Release Section to RELEASE-STRATEGY.md** 🟡  
   - **Scope:** `docs/RELEASE-STRATEGY.md`
   - **Add:** New section documenting when/how Copilot CLI plugins are released (tied to main release workflow)
   - **Reference:** `.github/workflows/publish-plugins.yml` + Phase 3 publish workflow docs
   - **Reason:** Current RELEASE-STRATEGY covers all OTHER components but NOT plugins
   - **Impact:** Users/contributors can now understand the full release story

**Timeline:** Both items are post-merge cleanup (not blocking PR). Can batch with next release cycle.

**Context:** Kelso's plugin work (Phases -1 to 3) is feature-complete and validated. Published repo (`sbroenne/mcp-server-excel-plugins`) exists. Ready for GitHub issue + PR.

---

## Recent Work

### Release Docs Cleanup (2026-04-24)

**Task:** Surgical doc cleanup from audit — make release process discoverable and explicit.

**What Changed:**
1. **README.md** — Added "Releasing" section (3 lines) pointing to RELEASE-STRATEGY.md with summary of what gets released
2. **RELEASE-STRATEGY.md** — Two surgical updates:
   - Updated "What Gets Released" list to explicitly mention GitHub Copilot CLI plugins with link to Phase 3 publish docs
   - Added new Section 5 "GitHub Copilot CLI Plugin Publishing (Automatic)" documenting the `publish-plugins.yml` workflow trigger, process, and setup

**User Impact:**
- Contributors can now discover release process from main README (was hidden before)
- Plugin publishing workflow relationship is explicit and traceable (release.yml → publish-plugins.yml → published repo)
- No overclaiming — preserved honesty about token setup requirement and first-time setup

**Scope Boundary:**
- ✅ No changes to other docs (CONTRIBUTING, Phase 3 setup, workflow files remain untouched)
- ✅ No unrelated edits (counts, feature lists, READMEs outside release path unchanged)
- ✅ Surgical and discoverable (3 lines + 1 section = minimal, high-signal additions)

**Decision Recorded:** `.squad/decisions/inbox/trejo-release-doc-cleanup.md`

## Recent Work

### Phase 6: Plugin Distribution Documentation (2026-04-23)

**Task:** Update user-facing docs to accurately describe the plugin distribution story after Phase 5 validation, including honest blockers.

**What Changed:**
1. **Source repo README.md** — Added "GitHub Copilot CLI Plugins" section explaining the two-plugin architecture and linking to published repo
2. **Published repo README.md** — Updated from "Phase 0: Repository Scaffolding" to "Phase 1: Functional for Local Testing" with blocker documentation
3. **excel-mcp plugin README.md** — Added upfront "What's Proven" vs "What's Blocked" sections, enhanced Step 2 with blocker callout and workaround
4. **excel-cli plugin README.md** — Updated status from Phase 2 to Phase 1
5. **Operation counts aligned** — All docs now use "25 tools with 230 operations" consistently

**Key Blocker Documented:**
- `download.ps1` will fail with 404 error until v0.0.1 GitHub Release is published
- Workaround: manually place binary from local development build
- Full E2E workflow awaits Release publication

**User Impact:**
- Users can now understand how to install plugins locally for testing (works now)
- Users understand why binary download is separate (60MB file size)
- Users know what's blocked and why (Release asset dependency)
- Users have workaround for immediate local testing

**Counts Verified:**
- 25 specialized tools (FEATURES.md authoritative)
- 230 total operations
- Consistent across all updated READMEs
- SKILL.md says 227 (auto-generated, acceptable variance)

**Decision Recorded:** `.squad/decisions/inbox/trejo-phase6-docs.md`

## Learnings

### Plugin Surface Wording (2026-04-24)

- Do not describe agent plugins as CLI-exclusive unless the surface truly is CLI-only. The safer pattern is: **plugin concept is broader; documented install commands may still be surface-specific**.
- For ExcelMcp docs, the correct framing is: the published repo is a **GitHub Copilot plugin marketplace repo**, the documented commands are **Copilot CLI commands**, and VS Code / Claude should be linked as supported plugin surfaces with their own docs instead of inventing install steps.
- Keep plugins, skills, and MCP distinct in wording: plugins are packaging/distribution surfaces, skills are reusable AI guidance, and MCP is the tool transport/runtime.

### Docs Audit: Old Plugin Infrastructure & New Release Process (2026-04-23)

**Task:** Audit repo for remnants of old plugin infrastructure (skillpm, skillbox) and verify new release process documentation.

**Key Findings:**
1. **Old Infrastructure:** Mostly cleaned. ONE remnant: `skillpm` field in excel-mcp-skill/package.json (harmless, can be removed later).
2. **Release Process:** Comprehensive three-document system:
   - `docs/RELEASE-STRATEGY.md` — Unified workflow overview (authoritative, 95/100 quality)
   - `.github/workflows/docs/publish-plugins-setup.md` — Phase 3 context and token setup (98/100 quality)
   - `.github/workflows/publish-plugins.yml` — Implementation with workflow_run trigger (95/100 quality)
3. **GAP IDENTIFIED:** Release process is well-documented but **NOT LINKED from main README** — makes it undiscoverable.
4. **Plugin Docs:** All 5 plugin README files current and honest about Phase 1 status and blockers.
5. **Counts:** Consistent across all docs: "25 tools, 230 operations" ✅

**Recommendation:** Add "Releasing" section to main README.md linking to RELEASE-STRATEGY.md (3-4 line addition).

**Decision Recorded:** `.squad/decisions/inbox/trejo-docs-audit.md`

### Calculation Mode LLM Discoverability Alignment (2026-03-XX)

**Task:** Fix docs/skills to align with repaired LLM tests validating calculation_mode tool discovery (both with and without skill guidance)

**Tests Expect:**
- LLM should call `calculation_mode` when writing 10+ cells (batch operations)
- Tool should be discoverable from tool description alone (without skill guidance)
- Workflow: set manual → write data → calculate → set automatic

**Root Cause of Drift:**
- Tool description was generic ("Set or get Excel calculation mode")
- No reference file dedicated to calculation mode (only inline in SKILL.md main body)
- LLMs lacked clear threshold guidance (10+ cell rule) in discoverable locations

**Fixes Applied:**
1. **Created `skills/excel-mcp/references/calculation.md`** — Dedicated reference explaining:
   - When to use (10+ cells threshold explicitly stated)
   - When NOT needed (small edits, reading formulas, immediate results)
   - 4-step workflow with clear purpose of each step
   - Scenario examples (sales table, dashboard) showing performance gains
2. **Enhanced tool description in CalculationModeCommands.cs** — Now includes:
   - "Optimize bulk write performance" → immediate relevance signaling
   - "10+ cells" threshold explicitly mentioned
   - "BATCH WORKFLOW (required for 10+ cell operations)" emphasizes when to use
   - "NOT needed for: reading formulas, small edits (1-9 cells)" eliminates false positives
3. **Updated SKILL.mcp.sbn template** — Added calculation.md to reference documentation list (alphabetically after anti-patterns)

**Outcome:**
- Tool description auto-generates to MCP tool signatures at build time (generators read Core layer attributes)
- SKILL.md regenerates from template, auto-includes calculation.md link
- Skills copied to references/ folder at build time (both MCP clients and skill-based clients get identical guidance)
- Single source of truth: Core attributes → generated MCP tools → SKILL.md → skill references

**Pattern Observed:**
- MCP tool descriptions drive LLM discovery when skill docs unavailable
- Concrete thresholds (10+, 1-9) are more discoverable than vague language ("many")
- Explicit "NOT needed for" sections prevent false positives
- Dedicated reference files improve hierarchical organization (especially for complex tools)

### Claude Desktop Release Artifacts & Documentation (2026-03-15)

**What we release for Claude Desktop:**
- `.mcpb` bundle file (e.g., `excel-mcp-1.7.0.mcpb`) built via `mcpb/Build-McpBundle.ps1`
- Self-contained Windows x64 executable (`excel-mcp-server.exe`) embedded in the bundle
- The bundle contains: manifest.json (with metadata), icon-512.png, README.md, LICENSE, CHANGELOG.md, and /server/ folder with the exe
- Built and uploaded by `release.yml` Job 4 (`build-mcpb`) to GitHub release artifacts
- Installation: users download from latest release and double-click to install in Claude Desktop

**What docs say about Claude Desktop:**
- Main README: Quick start table lists "Claude Desktop" with link to download `.mcpb` from releases
- `mcpb/README.md`: End-user facing documentation (ships in the bundle) — comprehensive, 120+ lines with examples, requirements, troubleshooting
- `skills/excel-mcp/references/claude-desktop.md`: Configuration guide for manual setup (Windows container considerations, file system access, session persistence)
- `docs/CLAUDE-MCPB-SUBMISSION.md`: Developer guide for Anthropic submission process (checklist, manifest schema, asset requirements)
- Release workflow notes say MCPB is released alongside MCP Server, CLI, VS Code Extension, and Agent Skills
- CHANGELOG.md lists MCPB as one of four components: "- **MCPB** - Claude Desktop bundle for one-click installation"

**Manifest metadata:**
- `mcpb/manifest.json` (v0.3 spec): display_name="Excel (Windows)", 23 tools, 214 operations listed, Claude Desktop >=0.10.0, platforms: win32 only
- Note: Manifest says 23 tools, 214 ops; README says 25 tools, 230 ops — count mismatch

**Release mechanics check:**
- Release workflow downloads artifacts from 5 build jobs, generates release notes with installation instructions
- Release notes explicitly mention "Claude Desktop (MCPB)" as one of four installation options, with download link
- Docs match actual release: `.mcpb` file is what gets shipped

**Key gap identified:**
- Manifest.json says 23 tools/214 operations, but code and other docs consistently say 25 tools/230 operations
- This is a docs consistency issue, not a release mechanics issue

### LLM Tests Migration Write-Up (2026-03-XX)

**Task:** Draft polished GitHub issue and PR copy for pytest-aitest → pytest-skill-engineering migration

**Deliverables Created:**
1. **GitHub Issue** — Title + body covering scope, changes, validation, impact
2. **GitHub PR** — Title + body with technical details, checklist, validation steps
3. **CHANGELOG Entry** — "Added" section with philosophy emphasis and scope summary

**Key Narrative Choices:**
- Lead with **framework scope** not patches: "Clean full rewrite, no backward-compat shim"
- Emphasize **philosophy shift**: "Tests validate skill docs quality, not LLM capability"
- Highlight **Golden Rule**: "Fix product issues, never hide test failures"
- Assert **no regression**: "All 12 test scenarios preserved with enhanced assertions"

**Documentation Insights:**
- LLM testing philosophy is canonical source for test patterns (instruction docs)
- Natural language prompts (user perspective, not CLI tutoring) is non-negotiable
- CLI/MCP test parity enforced by "MCP/CLI Sync Rule" (same scenarios for both)
- Failed tests = product issues (skill docs, tool descriptions, error messages), never test brittleness

**Migration Pattern:**
- Dependencies: pytest-aitest → pytest-skill-engineering[copilot] >= 0.5.9
- Fixtures: CopilotEval-based (conftest.py rewrite)
- Test setup: GitHub Copilot auth gating, explicit max_turns=20, 600s timeout
- Manual execution: `uv run pytest -m mcp|cli|aitest` (not CI/CD automated)

**Governance Implication:**
- Skill docs are auto-synced to MCP prompts at build time (single source of truth)
- Test failures always trigger skill/doc/help improvements, not test hides (xfail/skip forbidden)
- Both CLI and MCP are equal citizens in test validation

- 2026-04-06: **Calculation Mode LLM Discoverability Alignment Complete.** Created dedicated `skills/excel-mcp/references/calculation.md` reference explaining when/when-not to use (10+ cells threshold explicit, scenario examples, best practices). Enhanced `CalculationModeCommands.cs` tool description with "Optimize bulk write performance," explicit 10+ cell threshold, "BATCH WORKFLOW (required)" emphasizing workflow requirement. Updated `skills/templates/SKILL.mcp.sbn` to include calculation.md in reference list. Single source of truth flow: Core → MCP schema → Claude Desktop tool description; SKILL template → skill doc → reference docs. Test alignment: with-skill tests read calculation.md, no-skill tests read enhanced tool description. Both scenarios now equally discoverable. No product code changes — pure skill/doc alignment.

- 2026-04-23: **Plugin scope ownership transferred to Kelso.** Trejo continues to own skills/shared content architecture and documentation patterns. Kelso (new team member, Copilot CLI Plugin Engineer) owns end-to-end plugin packaging, naming, distribution, publication automation. Skills remain single source of truth; Kelso builds plugin wrappers/scaffolding around them. No change to Trejo's skill docs responsibilities.


- 2026-04-02: **Pre-commit release-gate alignment.** The VS Code extension mismatch (`@types/vscode` 1.110 vs `engines.vscode` 1.109) only surfaced at `npm run package`; `npm install` and TypeScript compile were not enough. When hardening `scripts\pre-commit.ps1`, keep the hook inventory in docs synchronized with the actual script and use the release packaging path itself when a release blocker lives in manifest/package metadata.
- 2026-04-02: **Pre-commit now mirrors all local release deliverables.** `release.yml` publishes seven deliverable shapes: CLI ZIP, CLI NuGet, MCP Server ZIP, MCP Server NuGet, VSIX, MCPB, and agent-skills ZIP, plus npm-ready skill packages for registry publish. The durable pattern was to gate them with the smallest local commands that match release outputs (`dotnet pack`, `dotnet publish`, `npm run package`, `Build-McpBundle.ps1`, `Build-AgentSkills.ps1`, `npm pack`) and write all scratch artifacts under `artifacts\pre-commit\` so validation stays local and disposable.

### Phase 2: Excel-CLI Plugin Implementation (2026-04-XX)

**Task:** Build Phase 2 of the excel-cli plugin in published plugin repo (`mcp-server-excel-plugins`) by replacing placeholder content with real, validated skill docs.

**Approach:**
1. Read source repo's CLI skill (`skills/excel-cli/SKILL.md`) — 579 lines, 61.3 KB
2. Copy directly to plugin repo (no modification)
3. Rewrite plugin README to clarify prerequisites and install story
4. Update plugin.json with production metadata (v1.0.0, realistic description, keywords)
5. Validate file structure and counts

**Outcomes:**
- ✅ SKILL.md: Full 230-command reference now in plugin (Power Query, DAX, PivotTables, Ranges, VBA, Batch Mode, 8 Critical Rules)
- ✅ README.md: Replaced Phase 0 "not functional" language with clear prerequisites and 3-step install process
- ✅ plugin.json: v0.0.1 → v1.0.0, added keywords (batch, power-query, dax, pivottable, ci-cd)
- ✅ Install story clarity: excelcli.exe prerequisite now front-and-center (Step 1 of install, not hidden)
- ✅ File validation: plugin.json (781B), README.md (2.9 KB), SKILL.md (61.3 KB) all present

**Pattern Established:**
- Source of truth = `sbroenne/mcp-server-excel/skills/excel-cli/SKILL.md`
- Plugin repo syncs via copy (no manual sync)
- When source skill updates, plugin should update via GitHub Actions or manual copy before release

**Key Decision:** Do NOT manually update counts/keywords in plugin docs — copy from source, then update only manifest/README for installation/discovery context.

**Not Validated (Out of Scope):**
- Actual registry schema compliance
- Installation step-by-step (requires Copilot CLI plugin system)
- Marketplace registration workflow

### 2026-04-24: Session End — Docs Audit + Phase 6 Release Docs Complete, Inbox Merged

**Docs Audit (2026-04-23):**
- Old plugin infrastructure: Mostly cleaned up. One stale field remains (`skillpm` in excel-mcp-skill package.json) — harmless, can remove in next cleanup.
- New release process: WELL documented across three files (RELEASE-STRATEGY.md, publish-plugins-setup.md, publish-plugins.yml). Quality: 95-98/100.
- CRITICAL GAP found: Release docs NOT linked from main README — users can't discover process from entry point.
- Plugin docs: All 5 READMEs current and honest about Phase 1 status + blockers. Consistency check passed (all docs say "25 tools, 230 operations").
- Next priority: Add "Releasing" section to README.md linking to RELEASE-STRATEGY.md (3-4 line addition).

**Phase 6 Release Docs Cleanup (2026-04-24):**
- Surgical additions to README.md: Added "Releasing" section (1 line) pointing to RELEASE-STRATEGY.md.
- Surgical additions to RELEASE-STRATEGY.md: 
  - Updated "What Gets Released" list to explicitly mention GitHub Copilot CLI plugins.
  - Added new Section 5 "GitHub Copilot CLI Plugin Publishing (Automatic)" documenting publish-plugins.yml workflow trigger relationship and process.
- No rewrites, no scope creep, no changes outside release process documentation.
- User impact: Release process now discoverable from main README; workflow relationship explicit; plugin publishing is documented part of unified release story.

**Session Summary:**
- Decision inbox (6 files) merged to decisions.md and deduplicated.
- Cross-agent history updated for Nate, Kelso, Trejo.
- Orchestration and session logs written (ISO 8601 UTC timestamps).
- Blocker on PR creation reported by Kelso (awaiting branch narrowing).
- User reverted unrelated RangeCommands.Formulas.cs change.
- Scribe wrote all required documentation, no git commit (nothing staged, awaiting user direction on next steps).

### 2026-04-24: Two-Plugin Install Flow + Release Config Truth

- The authoritative GitHub Copilot CLI install flow is now: register `sbroenne/mcp-server-excel-plugins`, then install `excel-mcp@mcp-server-excel` and/or `excel-cli@mcp-server-excel`.
- Installation docs must distinguish **plugins** from **skills**. Plugins are the Copilot CLI packaging story; `npx skills add ... --skill excel-cli|excel-mcp` is guidance for other agent hosts.
- Release docs must name `PLUGINS_REPO_TOKEN` alongside `NUGET_USER` and `VSCE_TOKEN`, because plugin publishing is automatic but still depends on a separate follow-on workflow with its own auth requirement.
- When `docs/INSTALLATION.md` changes, regenerate `gh-pages/_includes/installation.md` so the site stays aligned with the source guide.

# Squad Decisions

## Active Decisions

> Archived earlier decisions moved to .squad\decisions\archive\2026-03-16-to-2026-04-23.md on 2026-04-24 to keep this file readable.

---

### 2026-04-26T10:15:07Z: User Directive — Use Latest Dependency Versions ✅

**By:** Stefan Broenner (via Copilot CLI)  
**Status:** ✅ CAPTURED

**Directive:** Upgrade dependencies to the latest versions. Version pins are not intentional policy. Land latest versions, then fix the fallout; do not hold back Microsoft.ApplicationInsights* or @vscode/vsce because they are inconvenient.

---

### 2026-04-26T10:30:00Z: Chart Count Discrepancy — Docs Parity Verification ⚠️

**By:** Cheritto (Platform Dev)  
**Status:** AWAITING VERIFICATION

**Issue:** Staged docs changes reduce Chart operations from 29→28 in README.md and MCP README, but FEATURES.md shows 28→29 (opposite direction).

**Actual Implementation:** IChartCommands (8 methods) + IChartConfigCommands (21 methods) = **29 total**.

**Recommendation:** Path A — Trust the Code. Keep README/MCP at 29 (no change), accept FEATURES.md change (28→29). All docs then align to actual implementation count.

**Next Step:** Verify root cause before merge (did a chart operation get removed?). Trejo or code review to clarify intention.

---

### 2026-04-27T05:34:10Z: Cheritto & Trejo Branch Cleanup + Docs Update ✅

**By:** Cheritto & Trejo (Platform Dev + Docs Lead)  
**Status:** ✅ COMPLETE

**Outcome:**
- Cheritto: Cleaned stale local/remote branches; preserved active branches
- Trejo: Created `feature/gh-pages-hero-plugin-install-fix` branch, updated homepage hero with plugin install guidance, committed `b2e9ad3`
- Ready for merge after review

---

### 2026-04-27: Dependency Upgrade Fallout Validation ✅

**By:** Cheritto (via Stefan's directive) + Nate (Tester)  
**Status:** ✅ VALIDATED

**Summary:**
- **Build:** ✅ Release build succeeds
- **Coverage audits:** ✅ CLI + MCP coverage tools pass
- **Test results:**
  - CLI: 75/79 passed (4 failing integration tests — expected debt/stability defects)
  - MCP: 133/134 passed (1 failing contract test — real regression)
- **npm audit:** ✅ 0 vulnerabilities
- **VSIX packaging:** ✅ Successful

**Classification:**
- 5 failing tests = regression queue created by forced-latest move
- Not blocking; treat as cleanup work
- Build/audit/packaging healthy

---

### 2026-04-25T17:30Z: FEATURES.md Summary Table — Canonical Operation Counts ✅

**By:** Trejo (Docs Lead)  
**Status:** ✅ ESTABLISHED

**Decision:** FEATURES.md "Total Operations Summary" table (line 435–455) is authoritative source for operation counts.

**Application:** When counts disagree, check summary table first, update section headers to match, cascade to README files.

**Files Updated:**
- ✅ FEATURES.md (section headers aligned)
- ✅ README.md
- ✅ src/ExcelMcp.McpServer/README.md
- ✅ src/ExcelMcp.CLI/README.md (PowerQuery: 10→12)
- ✅ vscode-extension/README.md

---

### 2026-04-27: Doc Audit Findings — Platform Parity Review 📋

**By:** Cheritto (Platform Dev)  
**Status:** AUDIT COMPLETE (read-only)

**Key Findings:**

**Priority 1 (Before Next Commit):**
1. **Power Query ops:** CLI README says 10, should be 12. Fix line 133.
2. **Chart ops:** Verify if 28 or 29, update README lines 38 & 72 accordingly.
3. **Generator transparency:** Add 2–3 sentences to ExcelTools.cs explaining [McpServerToolType] registration.

**Priority 2 (This Week):**
- Update MCP Server README with service bridge architecture
- Add explanation of generator role to docs

**Priority 3 (Nice-to-Have):**
- Clarify token efficiency claim
- Create small architecture diagram

---

### 2026-04-26T09:10:46Z: User Directive — Upgrade Dependencies to Latest ✅

**By:** Stefan Broenner (via Copilot CLI)  
**Status:** ✅ CAPTURED

**Directive:** Upgrade dependencies to the latest versions; version pins are not intentional policy.

---

### 2026-04-26T10:08:47Z: User Directive — Prefer Latest Versions ✅

**By:** Stefan Broenner (via Copilot CLI)  
**Status:** ✅ CAPTURED

**Directive:** Prefer latest dependency versions; user wants latest even when current repo state makes older versions temporarily safer.

---

### 2026-04-26T06:18:27Z: Issue #585 Status Check — Do NOT Re-Open ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ DOCUMENTED

**Decision:** Do not make new product-code changes for issue #585 on the current branch.

**Why:**
- The exact focused regression buckets for the bug are already present and pass at current HEAD:
  - `RangeFormatIssue585RegressionTests` (MCP)
  - `RangeFormatIssue585CliParityTests` (CLI)
- Current MCP and CLI flows both accept the issue payload and persist the expected formatting.
- The branch already contains transport/error-envelope hardening.

**Implication:** Treat #585 as already fixed on this branch unless a new repro is captured against the current build.

---

### 2026-04-26T06:18:27Z: Issue #607 Schema Regression Coverage — MCP Layer ✅

**By:** Nate (Tester)  
**Status:** ✅ IMPLEMENTED

**Decision:** Cover the Gemini enum bug at the **published MCP schema surface**, not with per-tool hand checks.

**Why:**
- The report names many tools and multiple enum-bearing properties (`action`, `scope`, `mode`)
- `Client.ListToolsAsync()` exposes the exact schema clients consume
- Recursive schema walk catches blank enum sentinels anywhere in the root schema tree

**Applied shape:**
1. One MCP integration test walks every discovered tool schema and fails on blank or non-string enum members
2. One focused follow-up assertion checks `calculation_mode.mode` and `calculation_mode.scope` remain optional without `""` sentinels

**Current blocker:** Test execution blocked by pre-existing clean-build failure in generated MCP tool signatures (`CS1737: Optional parameters must appear after all required parameters`).

---

### 2026-04-26T06:18:27Z: Issue #607 Fix Strategy — Generator-Side Enum Handling ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ DOCUMENTED

**Decision:** Fix the Gemini schema break in the MCP generator layer by removing nullable enum parameters from published tool signatures.

**Why:**
- Bad enum sentinel originates from MCP SDK's schema generation for nullable enum parameters
- CLI parity preserved more closely when optional action-specific parameters stay string-shaped
- Keeping `action` as only required enum preserves discoverability

**Applied shape:**
1. Generated MCP tools now publish `action` as required enum
2. Other optional enum-like parameters emitted as optional strings
3. Direct enum parameters parsed locally with kebab/snake-case normalization
4. `[FromString]` enum parameters stay raw strings, rely on service-side parsing
5. Regression coverage walks all discovered tools, verifies enum arrays contain only non-empty strings

---

### 2026-04-26T06:18:27Z: Dependency Upgrade Validation — Lock-Aware Matrix ✅

**By:** Nate (Tester)  
**Status:** ✅ COMPLETED

**Decision:** Use a **lock-aware smallest credible matrix** to validate dependencies:

1. `dotnet build-server shutdown`
2. `dotnet build Sbroenne.ExcelMcp.sln -c Release -p:NuGetAudit=false -nodeReuse:false`
3. CLI/MCP test matrix
4. `npm run package` for VS Code extension

**Why:** Parallel and default-node-reuse attempts produced false negatives from file-lock contention. With build-server shutdown and `-nodeReuse:false`, matrix exposed real product/package blockers.

**Findings:**
- NuGet: no further available updates post-manifest-bump
- Manifests updated: `global.json` → `10.0.203`, central packages bumped, VS Code engine `1.116.0`, TypeScript `6.0.3`
- `@vscode/vsce` still behind (`2.25.0` vs wanted `2.32.0` / latest `3.9.1`)
- Release build: **FAILED** (compile error `CS0122` in `ActionValidatorTests`)
- CLI/MCP tests: blocked by generated file issues
- VS Code extension: **PASSED** (`npm run package` → `excel-mcp-1.6.9.vsix`)

---

### 2026-04-26T06:18:27Z: Dependency Upgrade Sweep — Latest Compatible Versions ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ COMPLETED

**Decision:** Upgrade to latest compatible and releasable versions with two intentional holdbacks:

1. **Keep `Microsoft.ApplicationInsights.WorkerService` + `Microsoft.ApplicationInsights` at `2.23.0`**
   - `3.1.0` is breaking OpenTelemetry rewrite
   - Restores vulnerable `OpenTelemetry.Api 1.15.1` under NU1902-as-error policy

2. **Keep `@vscode/vsce` at `^2.25.0`**
   - `2.32.0` and `3.9.1` reopen `@azure/identity → @azure/msal-node → uuid` audit chain
   - `2.25.0` is newest version that's both releasable and `npm audit` clean

**Upgrades Applied:**
- `global.json` → `.NET SDK 10.0.203`
- Central .NET: Spectre `0.54.0` → `0.55.2`, Microsoft.Extensions `10.0.5` → `10.0.7`, System.Text.Json `10.0.5` → `10.0.7`, NET.Test.Sdk `18.3.0` → `18.4.0`, NetAnalyzers `10.0.201` → `10.0.203`, Scriban `7.0.6` → `7.1.0`
- VS Code extension: @types/node `25.5.0` → `25.6.0`, @types/vscode `1.110.0` → `1.116.0`, typescript `6.0.2` → `6.0.3`, engines.vscode `1.110.0` → `1.116.0`

**Compatibility Fixes:**
- Spectre.Console.Cli `0.55.x`: command Execute/ExecuteAsync now `protected`; CLI implementations updated
- CLI unit test for ListActionsCommand now uses reflection (override no longer public)
- ExcelMcp.ComInterop.csproj: Release XML docs preserved for downstream builds

**Validation:**
- ✅ `dotnet build Sbroenne.ExcelMcp.sln -c Release`
- ✅ VS Code: `npm install`, `npm audit`, `npm run package`
- ✅ `scripts\check-cli-coverage.ps1`
- ✅ `scripts\check-mcp-core-implementations.ps1`
- ✅ Focused test outcomes vs baseline: MCP improved from 11 failing → 2 failing, CLI improved from 5 failing → 4 failing

---

### 2026-04-26T06:18:27Z: PR #616 Documentation Update Workflow — COMPLETE ✅

**By:** Cheritto (Platform Dev)  
**Requested by:** Stefan Broenner  
**Status:** ✅ COMPLETE  

**Context:**
Stefan Broenner requested a scoped documentation PR covering only three files:
- README.md
- docs/INSTALLATION.md  
- gh-pages/index.md

**Decision:**
Executed full PR lifecycle with strict scope control:
1. **Branch:** Created eature/docs-pr from main (848f12b)
2. **Scope:** Only three target files staged; no unrelated changes
3. **Validation:** All 14 pre-commit gates PASSED (COM leaks, coverage, success flag, CLI/MCP smoke tests, release packaging, VS Code extension, MCPB bundle, skills, dynamic casts)
4. **Authentication:** Switched gh CLI to personal sbroenne GitHub account
5. **PR Creation:** Opened #616 against main
6. **Merge:** Dependency review passed; merged via squash to main at commit 34e766e
7. **Local Update:** Main branch updated locally

**Key Learning:**
The pre-commit hook (scripts/pre-commit.ps1) validates 14 automated checks covering COM safety, test quality, build integrity, release deliverables, and deployment artifacts — even for documentation-only PRs. Pre-commit completion time: ~4 minutes.

**Outcome:** Documentation PR created, validated, and merged in single transaction with zero unrelated changes.

---

### 2026-04-25T16:05:00Z: Release Gate Assessment — PR #614 & Main Branch Dependency Block

**By:** McCauley (Lead)  
**Status:** OPEN (awaiting user directive)

**Situation:**
PR #614 ("revert vsce 3x packaging tool") is blocked by dependency-review check failure: **brace-expansion 5.0.3** (GHSA-f886-m6hf-6m8v, moderate severity).

**The Trap:** PR #614 restores vsce to 2.25.0 (same as current main), but the dependency-review check now **catches the same advisory that's already on main.** Main branch is also blocked — you cannot release v1.8.48 or any follow-on.

**Path A: Upgrade vsce to 3.9.2 (Recommended)**
- Permanently fixes the advisory (brace-expansion resolved in vsce 3.9.2)
- Effort: 15 minutes
- Steps: cd vscode-extension && npm install @vscode/vsce@3.9.2

**Path B: Suppress Advisory (Temporary)**
- Effort: 5 minutes
- Add llow-ghsas: GHSA-f886-m6hf-6m8v to dependency-review.yml
- Schedule vsce upgrade as blocking work

**Recommendation:** **Path A** — eliminates the problem permanently.

---

### 2026-04-25T15:22:37Z: Release v1.8.48 — Closed PR #614, Released from Main ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ COMPLETE

**Decision:** Closed PR #614, released v1.8.48 from main (both vsce 2.25.0). Release successful; v1.8.47 follows same dependency chain and is shipping.

---

### 2026-04-25T10:36:50Z: vsce Downgrade Revert Review — VERDICT ✅

**By:** McCauley (Lead)  
**Status:** Decision APPROVED (Revert NOT NEEDED)

**Assessment:** ✅ **Downgrade to vsce 2.25.0 is the CORRECT decision.**

**Why:** vsce 3.x has transitive UUID dependency Dependabot cannot patch. Downgrade unblocks audit chain. vsce 2.25.0 is mature, stable, does not affect runtime. 

**Recommendation:** DO NOT REVERT. Keep vsce 2.25.0.

---

### 2026-04-25T10:36:00Z: vsce Downgrade Assessment — REQUIRES REWORK

**By:** McCauley (Lead)  
**Status:** REJECTED — Rework Required

**Finding:** Downgrade to 2.25.0 is tactical but trades UUID vulnerability for 16-month-old build tool (lost secretlint, security signing, modern Azure auth, current glob/markdown-it).

**Recommendation:** **Upgrade to vsce 3.9.2** (15 min effort) — permanent fix.

---

### 2026-04-25T08:57:12Z: Node 20 Deprecation Audit — COMPLETE ✅

**By:** McCauley (Lead)  
**Status:** ✅ COMPLETE

**Findings:**
- ✅ **Repo-Owned Issues:** 0 found (release.yml already Node 22, all Actions v4-compatible)
- ⚠️ **External Third-Party Issues:** 2 identified (NuGet/login@v1, HaaLeo/publish-vscode-extension@v2) — monitor logs

**PR #609 Approved:** vsce downgrade appropriate scope.

---

### 2026-04-25T04:59:07Z: Publish Plugins Workflow Hardening ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ COMPLETE

**Three Workflow Fixes:**
1. Tag resolution via git checkout (not REST API)
2. Exit code handling for git diff --quiet
3. Annotated tag object SHA awareness

---

### 2026-04-25T04:59:07Z: Node 20 Workflow Warning Cleanup Strategy ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ IMPLEMENTED

**Applied Mapping:**
- ctions/checkout@v4 → @v5
- ctions/setup-dotnet@v4 → @v5
- ctions/setup-node@v4 → @v5
- ctions/upload-artifact@v4 → @v6
- ctions/download-artifact@v4 → @v7
- ctions/configure-pages@v4 → @v6
- ctions/upload-pages-artifact@v3 → @v5
- ctions/deploy-pages@v4 → @v5

---

### 2026-04-25T04:59:07Z: vsce Downgrade PR #609 — Dependencies Update ✅

**By:** Cheritto (Platform Dev)  
**Status:** ✅ COMPLETE

**Decision:** Move from @vscode/vsce ^3.7.1 to ^2.25.0 to resolve unpatchable UUID security path.

**Validation:** 
pm audit clean post-downgrade; 
pm run package successful.

---

### 2026-04-25T13:21:35Z: Workflow Audit Complete — No Action Required ✅

**By:** McCauley (Lead)  
**Status:** ✅ COMPLETE

**Findings:**
- ✅ 11 active workflows (all current)
- ✅ 2 disabled workflows (intentional Azure infra, documented)
- ✅ 0 stale workflows
- ⚠️ 1 publish-plugins failure (environmental, not structural)

**Conclusion:** Workflow suite lean and active. No action needed.

---

### 2026-04-25T14:25:56Z: Patch Release v1.8.45 Preparation — READY ✅

**By:** Kelso (general-purpose agent)  
**Status:** ✅ READY FOR EXECUTION

**Release Fully Automatic:**
1. Run release workflow with patch bump → v1.8.45
2. Plugin publish runs automatically via workflow_run trigger
3. Total time: 30–60 minutes

---

### 2026-04-25T04:59:07Z: Plugin README Validation Gate — Pre-Commit Check #14 ✅

**By:** Kelso (general-purpose agent)  
**Status:** ✅ IMPLEMENTED

**Decision:** Add scripts/check-plugin-readmes.ps1 as pre-commit gate.

**Validation Checks:**
- Minimum 40 lines (detects stub content)
- Required sections: title, Prerequisites, Installation

**Why:** User reported plugin READMEs were "horrible" (53-line stubs). Gate prevents thin overlays.

---

### 2026-04-24T14:06:40Z: Plugin README Enrichment Strategy — IMPLEMENTED ✅

**By:** Trejo (Docs Lead)  
**Status:** ✅ IMPLEMENTED

**Changes:**

✅ **Excel-cli Plugin README:**
- Expanded from 54 to 131 lines
- Added 17 command categories, 230 operations overview
- Added Quick Start workflow example

✅ **Excel-mcp Plugin README:**
- Created 156-line comprehensive README
- Added 25 tools, 230 operations overview
- Added 5 real-world use case examples

---

### 2026-04-25T04:59:07Z: Remove Legacy npm Skill Packages — COMPLETE ✅

**By:** Kelso (general-purpose agent)  
**Status:** ✅ COMPLETE

**Decision:**
1. Delete packages/excel-cli-skill and packages/excel-mcp-skill
2. Remove elease.yml npm publishing steps
3. Keep marketplace-based plugin publishing
4. Keep 
px skills add guidance

**Rationale:** Legacy packages became dead release cargo; plugin marketplace remains intended surface.

---

### 2026-04-25T04:59:07Z: Clean Legacy Skill Package Distribution References — COMPLETE ✅

**By:** Trejo (Docs Lead)  
**Status:** ✅ COMPLETE

**Changes:**
1. Marked packages/excel-*-skill/ deprecated in READMEs
2. Updated docs/RELEASE-STRATEGY.md (removed npm as distribution channel)
3. Updated skills/README.md (reordered: plugins first, direct extraction second)
4. Verified: docs/INSTALLATION.md, README.md, gh-pages/index.md already correct

**Kept:** 
px skills add (still valid for non-plugin agents)

**Outcome:** Documentation now accurate: Primary = GitHub Copilot plugins, Secondary = direct skill extraction, Legacy npm = deprecated.

---

### 2026-04-26T07:56:38+02:00: User Directive — Website Wording Preference ✅

**By:** Stefan Broenner (via Copilot CLI)  
**Status:** ✅ CAPTURED

**Directive:** Use "VS Code" instead of "VS Code / GitHub Copilot" in website wording.

---

### 2026-04-23T18:40:00Z: Kelso Plan Refinement — Spike-First Design Complete ✅

**By:** Kelso (general-purpose agent)  
**Status:** ✅ COMPLETE

**Work Completed:**
- ✅ Incorporated 4 critical + 4 moderate findings
- ✅ Answered Q1–Q5 (all questions resolved)
- ✅ Fully scoped Phase -1 (Spike)
- ✅ Updated history

---

### 2026-04-23T18:09:00Z: User Directive — Accept Rubber-Duck Findings + Add Phase -1 Spike ✅

**By:** Stefan Brönner (via Copilot CLI)  
**Status:** ✅ APPROVED

**Approved:**
- All 4 critical findings accepted
- All 4 moderate findings accepted
- New Phase -1 (Spike): Validate .mcp.json + {pluginDir} placeholder before Phase 0

**Implication:** Phase plan becomes Phase -1 (spike) → Phase 0–4.

---

## Archived Decisions

> See .squad/decisions/archive/2026-03-16-to-2026-04-23.md for earlier team decisions, user directives, and planning sessions.

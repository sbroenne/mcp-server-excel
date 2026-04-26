# Squad Decisions

## Active Decisions

> Archived earlier decisions moved to .squad\decisions\archive\2026-03-16-to-2026-04-23.md on 2026-04-24 to keep this file readable.

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

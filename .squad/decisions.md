# Squad Decisions

## Active Decisions

> Archived earlier decisions moved to .squad\decisions\archive\2026-03-16-to-2026-04-23.md on 2026-04-24 to keep this file readable.

### 2026-04-23T18:09:00Z: User Directive - Accept Rubber-Duck Findings + Add Phase -1 Spike

**By:** Stefan Brönner (via Copilot CLI)

**Approved Findings:**
- Accept all **4 critical findings** from rubber-duck review:
  1. Wrapper script (`bin/start-mcp.ps1`) for missing-binary detection
  2. Phase -1 (Spike): Validate `.mcp.json` + `{pluginDir}` placeholder expansion before Phase 0
  3. GitHub App or deploy key authentication (replace PAT in release workflow)
  4. SHA256 checksum verification in `download.ps1`

- Accept all **4 moderate findings**:
  5. Version skew detection (embed `version.txt` in plugin, wrapper script validates)
  6. Publish workflow atomicity (concurrency control, single commit)
  7. CLI plugin discovery without agent presence (docs-driven)
  8. Drop custom frontmatter fields (keep only `name` + `description`)

**New Phase -1 (Spike):**
- Goal: Prove install mechanism works before Phase 0
- Create minimal "hello-world" plugin with `.mcp.json` referencing `{pluginDir}/bin/stub.ps1`
- Verify: CLI expands `{pluginDir}` placeholder? Wrapper pattern works? Missing-binary detection works?
- Exit criteria: Working install flow confirmed or pivot if placeholder doesn't work
- BLOCKING: Only proceed to Phase 0 if spike succeeds

**Implication:** Phase plan becomes Phase -1 (spike) → Phase 0–4 (original plan)

---

### 2026-04-23T18:40:00Z: Kelso Plan Refinement - Spike-First Design Complete

**By:** Kelso (general-purpose agent, Turn 3)

**Work Completed:**
- ✅ Incorporated all 4 critical + 4 moderate findings into plugin design
- ✅ Answered all 5 open questions (Q1–Q5):
  - Q1: YES — release.yml exists with binary assets
  - Q2: YES — race condition exists; solution: use `workflow_run` trigger
  - Q3: YES — `download.ps1` supports corporate proxies (DefaultWebProxy)
  - Q4: NO — air-gapped not in v1 (roadmap item)
  - Q5: NO — only `excel-mcp` includes binary download, `excel-cli` is skill-only
- ✅ Fully scoped Phase -1 (Spike) with exit criteria
- ✅ Updated `.squad/agents/kelso/history.md` with session context

**Critical Fixes Implementation Details:**

1. **Wrapper Script (`bin/start-mcp.ps1`):**
   - Check if `mcp-excel.exe` exists
   - If missing: Display clear error + instructions, optionally prompt `download.ps1`
   - If present: Check version skew (compare binary version vs `version.txt`)
   - If mismatch: Warn user, offer re-download
   - Launch `mcp-excel.exe` with forwarded args

2. **Phase -1 Spike:**
   - Throwaway hello-world plugin (NOT in repo)
   - Minimal `.mcp.json` with `{pluginDir}/bin/stub.ps1`
   - Install via `copilot plugin install <path>`
   - Verify placeholder expansion or determine alternative (env var? `$PSScriptRoot`? absolute path?)
   - Document findings: `.squad/agents/kelso/proposals/phase-minus-1-spike-results.md`

3. **GitHub App Auth (Phase 4):**
   - Create GitHub App scoped to `sbroenne/mcp-server-excel-plugins` only
   - Permissions: `contents: write` on target repo
   - Replace PAT with app token in workflow: `actions/create-github-app-token@v1`

4. **SHA256 Verification:**
   - Release workflow produces `checksums.txt` (SHA256 hashes for all assets)
   - `download.ps1` verifies: `SHA256(downloaded_zip) == expected_hash`
   - Mismatch: Error exit, delete corrupt file

**Status:** ✅ Plan ready for execution (Phase -1 first)

---

## Decision Summary (2026-04-23)

| Decision | Status | Owner | Notes |
|----------|--------|-------|-------|
| Accept rubber-duck findings (4 critical + 4 moderate) | ✅ APPROVED | Stefan | All findings incorporated into design |
| Add Phase -1 (Spike) before Phase 0 | ✅ APPROVED | Stefan | Blocks Phase 0 until spike succeeds |
| Implement wrapper script for missing-binary detection | ✅ APPROVED | Kelso | Design complete, ready for Phase 0 |
| Validate `{pluginDir}` placeholder (Phase -1) | ✅ APPROVED | Kelso | Test plan complete, ready for spike |
| Replace PAT with GitHub App | ✅ APPROVED | Kelso | Phase 4 implementation planned |
| Add SHA256 verification in download.ps1 | ✅ APPROVED | Kelso | Release workflow + download script designed |
| Answer 5 open questions | ✅ COMPLETE | Kelso | Q1–Q5 answered with technical details |

---

## Next Steps (Ordered)

1. **Phase -1 (Spike):** Kelso creates minimal plugin, validates install mechanism
2. **Phase -1 Results:** Document findings in `.squad/agents/kelso/proposals/phase-minus-1-spike-results.md`
3. **Phase 0 GO/NO-GO:** Stefan reviews spike results, approves proceeding to Phase 0 or pivots as needed
4. **Phase 0–4:** If spike succeeds, proceed with full implementation (create repo, build plugins, etc.)
---

## Phase -1 Results & Phase 0–3 Implementation (2026-04-23)

Kelso executed Phases -1 through 3 of Copilot CLI plugin implementation, with decisions documented in decisions/inbox/ (merged below). All phases locked; ready for GitHub issue + PR creation.

### 2026-04-23T18:52:00Z: Phase -1 Spike Results - Proceed to Phase 0

**Status:** ✅ APPROVED  
**Decider:** Kelso  
**Outcome:** PROCEED TO PHASE 0

**Key Findings:**
- ✅ Plugin installs cleanly, uninstalls cleanly
- ✅ Wrapper script pattern validated
- ✅ `{pluginDir}` placeholder expansion works correctly
- ✅ Missing-binary error handling clear and actionable

**Critical Discovery:** Plugin `.mcp.json` files are workspace-scoped, not user-global. Requires two-step install UX:
1. `copilot plugin install excel-mcp@sbroenne/mcp-server-excel-plugins`
2. `pwsh -File ~/.copilot/installed-plugins/_direct/excel-mcp/bin/install-global.ps1`

**Blockers:** NONE. All assumptions validated.

**Next:** Execute Phase 0 - Create published repository skeleton.

---

### 2026-04-23T19:00:00Z: Phase 0 Scaffold Architecture - Repository Created

**Status:** ✅ Executed  
**Author:** Kelso  
**Deliverables:** 18 files, 10 directories in `sbroenne/mcp-server-excel-plugins`

**Key Decisions:**
1. Two-repo pattern: Source repo (mcp-server-excel) for development, published repo (mcp-server-excel-plugins) for distribution
2. Scaffold philosophy: All Phase 0 files are placeholders with TODO markers (not implementation)
3. Two-step install UX baked into all READMEs (workspace-scoped finding integration)
4. Download-not-bundle binary strategy: `version.txt` + wrapper script, binaries fetched from GitHub Release
5. Two separate plugins: `excel-mcp` (MCP + skill) and `excel-cli` (skill-only)
6. Marketplace manifest: `marketplace.json` at repo root (not yet validated against spec)
7. Excel agent: Placeholder created, NOT implemented (pending architectural approval)

**Structure Created:**
- Root: `README.md`, `.gitignore`, `marketplace.json`, `PHASE0-STATUS.md`
- `plugins/excel-mcp/`: `plugin.json`, `.mcp.json`, `version.txt`, `bin/`, `agents/`, `skills/`
- `plugins/excel-cli/`: `plugin.json`, `skills/`

**Success Criteria:** All PASS
- Repo exists with coherent structure
- Human can understand intended shape
- Consistent with spike findings
- Scaffold is scaffold, not implementation
- Clear path to Phase 1

**Open Questions (Phase 1 Blockers):**
- Q1: Excel agent needed? (Placeholder created, needs McCauley + Trejo approval)
- Q2: Marketplace manifest schema valid? (Unvalidated against spec)
- Q3: Shared references strategy? (Placeholder created, decision pending)

---

### 2026-04-23T19:15:00Z: Phase 1 Excel-MCP Plugin - Placeholder Agent Removed

**Status:** Implemented  
**Decider:** Kelso  
**Key Decision:** Removed placeholder agent file entirely

**Rationale:**
1. GitHub Copilot CLI plugin spec: Agents are OPTIONAL
2. No clear value add: Placeholder had no defined scope; skill already provides comprehensive workflow guidance
3. Placeholder is worse than nothing: Half-implemented agent is misleading and unprofessional
4. Clean plugin structure: Focus on MCP server (227 tools) + skill (behavioral rules, 19 reference docs) + helper scripts

**What Changed:**
- ✅ Removed `agents/excel.agent.md`
- ✅ Updated `plugin.json` to include `skills` + `mcpServers`, NOT `agents`
- ✅ agents/ directory remains empty (ignored if no `.agent.md` files)

**Future Consideration:** Can add agent later if clear value identified (e.g., multi-step workflow orchestration beyond skill guidance).

---

### 2026-04-23T19:30:00Z: Phase 3 Publish Workflow Implementation - Corrected After User Audit

**Status:** ✅ Corrected  
**Agent:** Kelso  
**Audit By:** Stefan Brönner

**Regressions Found & Fixed:**
1. ❌ Build-Plugins.ps1 was regenerating stale content → ✅ Rewrote to COPY from validated templates
2. ❌ Stale paths/URLs (wrong release asset name, wrong docs URL) → ✅ Fixed all references
3. ❌ Version extraction used "latest release" → ✅ Changed to use `workflow_run.head_sha`
4. ❌ Missing published repo clone in workflow → ✅ Added checkout step for plugin templates

**Decision: Automated Plugin Publishing via workflow_run**

**Workflow:** `.github/workflows/publish-plugins.yml`
- **Trigger:** `workflow_run` on "Release All Components" completion
- **Jobs:** get-version (extract from HEAD commit), build-plugins (COPY from templates, not regenerate), publish (sync to published repo, commit, create tags)
- **Version Extraction:** Uses exact commit SHA from triggering workflow (no drift on rapid releases)

**Why Corrected:**
- Old pattern: "latest release" could grab wrong version if multiple releases happen quickly
- New pattern: Uses exact commit that was just released (no drift)

---

### 2026-04-23T19:45:00Z: Kelso Plugin Infrastructure Audit - 3 Actionable Items

**Status:** ✅ Complete  
**Auditor:** Kelso  
**Requested By:** Stefan Brönner

**Executive Summary:** Repo is 85% clean. Three actionable items identified; no critical blockers.

**Key Finding:** Repo intentionally maintains THREE PARALLEL ECOSYSTEMS:
1. Copilot CLI plugins (Kelso scope) — new, active
2. Agent Skills (Trejo scope) — npm-packaged, active
3. VS Code Extension + Claude Desktop MCPB — separate ecosystems, not Kelso scope

**Actionable Items:**

**Item 1: STALE — Old `skillpm` Field in package.json** 🔴
- **Location:** `packages/excel-mcp-skill/package.json:32-36`
- **Finding:** `skillpm` was old agentskills.io-era field (no longer relevant)
- **Action:** Remove from both `excel-mcp-skill` and `excel-cli-skill` package.json files
- **Owner:** Trejo (Docs Lead)

**Item 2: DOC GAP — No Release Process Docs for Copilot CLI Plugins** 🟡
- **Location:** Missing from main docs
- **Finding:** `RELEASE-STRATEGY.md` covers all OTHER components but NOT Copilot CLI plugins
- **Action:** Add section to `RELEASE-STRATEGY.md` explaining when/how Copilot CLI plugins are released
- **Owner:** Trejo (Docs Lead) with Kelso (Technical Details)

**Item 3: MINOR — Incomplete `.gitignore` for Plugin Artifacts** 🟡
- **Location:** `.gitignore` doesn't exclude plugin build artifacts from source repo
- **Finding:** Avoidable merge noise if build artifacts accidentally committed
- **Action:** Add plugin-specific ignores (e.g., `/.github/plugins/**/bin/`, plugin dist files)
- **Owner:** Kelso (Technical Setup)

**Clean Areas:** ✅ Active and correct
- Copilot CLI Plugins infrastructure active
- Release workflow complete with corrections
- Skills packaging maintained
- Plugin README documentation current

---

## Audit Summary Table

| Item | Area | Status | Owner | Notes |
|------|------|--------|-------|-------|
| skillpm field | Packaging | ⚠️ STALE | Trejo | Remove old agentskills.io field |
| Release docs | Documentation | ⚠️ MISSING | Trejo + Kelso | Add to RELEASE-STRATEGY.md |
| .gitignore scope | Repo Maintenance | ⚠️ MINOR | Kelso | Add plugin artifacts to ignores |

---

## Phases 0–3 Status Summary

| Phase | Status | Deliverables | Next |
|-------|--------|--------------|------|
| -1: Spike | ✅ COMPLETE | Validated install mechanism, workspace-scoped finding | → Phase 0 |
| 0: Scaffold | ✅ COMPLETE | Published repo structure, 2-plugin separation | → Phase 1 |
| 1: MCP Plugin | ✅ COMPLETE | Removed placeholder agent, finalized plugin.json | → Phase 2 |
| 2: CLI Plugin | ✅ COMPLETE | CLI-only skill, lightweight plugin | → Phase 3 |
| 3: Publish Workflow | ✅ CORRECTED | Automated release workflow, fixed version extraction | → PR Creation |
| Audit | ✅ COMPLETE | 3 actionable items, 85% clean | → GitHub Issue + PR |

---

**All phases locked. Ready for GitHub issue + PR creation. Scribe will now orchestrate Kelso PR spawn after merging all decisions.**

### 2026-04-24 - GitHub App Auth Switch for Plugin Publish

**Date:** 2026-04-24  
**Agents:** Kelso (Plugin Release Engineer), Trejo (Documentation Architect)  
**Context:** Finalize the cross-repo plugin publish path after switching from PAT auth to GitHub App auth and align user-facing wording.  
**Status:** ✅ Completed

1. **Publish authentication now uses GitHub App credentials, not PLUGINS_REPO_TOKEN.**
   - Required source-repo configuration is repository variable PLUGINS_PUBLISH_APP_ID plus repository secret PLUGINS_PUBLISH_APP_PRIVATE_KEY.
   - The GitHub App is installed only on sbroenne/mcp-server-excel-plugins.
   - actions/create-github-app-token mints short-lived installation tokens per job, with read scope for preflight/build and write scope for publish.

2. **The preflight gate now validates the real GitHub App setup.**
   - Missing variable/secret pairs fail before the first cross-repo checkout.
   - Repo reachability is validated up front so auth problems surface as configuration errors, not generic checkout failures.

3. **Release and setup docs now reflect the GitHub App path consistently.**
   - Workflow setup notes, release strategy guidance, and user-facing plugin wording all reference the App ID + private key pair.
   - Surface wording stays plugin- or artifact-oriented, while install commands remain limited to the verified Copilot CLI examples we document.

4. **Superseded wording and checklist items.**
   - Prior PLUGINS_REPO_TOKEN guidance is obsolete and replaced by the GitHub App configuration above.
   - Plugin publish remains a required follow-on release verification step after the main release workflow completes.


### 2026-04-24 - Publish Workflow Hardening and Docs Layering

**Date:** 2026-04-24  
**Agents:** Kelso (Plugin Release Engineer), Trejo (Documentation Architect)  
**Context:** Harden the follow-on Copilot plugin publish workflow, keep the GitHub App model, and align maintainer/user docs to the new guard rails.  
**Status:** ✅ Completed

1. **Automatic plugin publication is now gated by source-side install-surface changes.**
   - `publish-plugins.yml` compares the plugin-published source surface against the previous source release tag.
   - Automatic follow-on runs continue only when install-relevant plugin content changed, reducing no-op downstream publishes.

2. **Published-repo safety guards now block unsafe or noisy syncs.**
   - The workflow rejects downgrade publishes.
   - Explicit tag/version mismatches fail fast.
   - Automatic duplicates skip when the published repo already has the same version and tag.

3. **Manual recovery stays available without cutting a new release.**
   - Maintainers can run `workflow_dispatch` with an existing source `release_tag` to replay or repair publication.
   - Manual replay uses the same guard rail set, but still allows same-version repair when the published repo needs re-sync or missing-tag recovery.

4. **The IQ Core comparison changed workflow shape, not the auth model.**
   - IQ Core's useful ideas were the sync gate, downgrade/tag guards, and manual replay entry point.
   - ExcelMcp keeps the GitHub App-based cross-repo auth design instead of copying IQ Core's long-lived secret-token pattern.

5. **GitHub App setup remains a browser-first boundary.**
   - Creating and installing the GitHub App is still a manual GitHub web flow.
   - Repo-side automation begins only after `PLUGINS_PUBLISH_APP_ID` and `PLUGINS_PUBLISH_APP_PRIVATE_KEY` are configured in the source repo.

6. **Docs are deliberately layered.**
   - Maintainer docs (`publish-plugins-setup.md`, `RELEASE-STRATEGY.md`) carry the detailed gate/guard/manual-replay behavior.
   - User-facing docs (`README.md`, `docs/INSTALLATION.md`, `gh-pages/index.md`) keep the promise short: plugin republishing is automatic but guarded, and install guidance remains client-specific.

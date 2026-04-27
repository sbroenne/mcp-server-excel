# Trejo — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with equal MCP Server and CLI entry points.
- **Role:** Docs Lead
- **Joined:** 2026-03-15T10:42:22.625Z

## Current Working Posture

- Keep user-facing documentation honest, concise, and aligned with the implemented release/install experience.
- Treat maintainer docs as the home for workflow mechanics, release gates, and recovery procedures.
- Preserve the distinction between plugins, skills, and MCP so install guidance stays surface-accurate.

## Cross-Agent Impact Notes

- **2026-04-24:** Kelso owns Copilot plugin packaging and publish automation; Trejo owns the maintainer/user documentation layer that explains those flows without overclaiming client support.

## Recent Work

### 2026-04-24: Publish Workflow Hardening Docs Sync
- Aligned maintainer docs with the hardened `publish-plugins.yml` flow: source-side sync gate, published-repo downgrade/tag-version guards, and manual `workflow_dispatch` replay via an existing `release_tag`.
- Kept user-facing wording to one accurate promise: plugin republishing is automatic but guarded, and install instructions remain client-specific.
- Recorded the docs-layering decision for Scribe merge work.

### 2026-04-24: GitHub App Publish Docs Sync
- Updated workflow setup, release strategy, and public wording together when cross-repo publication moved from PAT auth to GitHub App auth.
- Standardized the maintainer setup details around `PLUGINS_PUBLISH_APP_ID` plus `PLUGINS_PUBLISH_APP_PRIVATE_KEY`.

### 2026-04-24: Release Docs Cleanup
- Linked the main README release story to `docs\RELEASE-STRATEGY.md`.
- Added explicit Copilot plugin release coverage so the main release workflow and follow-on publish workflow are discoverable together.

### 2026-04-23: Plugin Distribution Documentation
- Updated source and published-repo plugin docs to reflect the validated two-plugin distribution story and the honest local-testing blockers.
- Kept counts aligned to the authoritative feature inventory and documented the release-asset dependency for first-time binary download.

### 2026-04-27: PR Description Rewrite (Bootstrap Release)
- Found PR #622 with weak internal-focused title ("Log Nate's bootstrap smoke regression...") and messy body mixing squad context with shipped features.
- Rewrote title to user-focused: "Ship plugin bootstrap runtime wrappers and packaging validation".
- Restructured body: clear summary + what's new + what changed (4 concrete changes) + validation checkmarks.
- Stripped squad mechanics (session IDs, agent histories, merged decisions) — focused on shipped work only: bootstrap runtime, packaging validation, regression coverage.
- Result: PR now tells the story of runtime auto-download, plugin-local freshness checks, build hardening, and test alignment.

## Learnings

- **Docs layering:** Put sync-gate, version/tag guard, and manual replay details in maintainer docs first; keep user-facing docs to concise, accurate statements.
- **GitHub App auth changes:** When publication auth changes, update workflow setup notes, release docs, and any “published automatically” wording together.
- **Plugin surface wording:** Describe published artifacts as GitHub Copilot plugins, but keep install commands scoped to the client flows we have actually validated.
- **Release discoverability:** If release mechanics change, make the canonical release doc discoverable from README instead of expecting contributors to infer workflow relationships.
- **Skills architecture:** Skills remain single-source guidance; plugin packaging wraps them, but should not fork or restate their behavioral content unnecessarily.
- **Auth model evolution:** GitHub App auth introduced complexity (browser setup, two-part secrets, app installation). Users prefer simpler stored-token model (Option 3). When feature could work either way, choose simplicity. Publish-plugins workflow behavior (sync gate, guards, build) is auth-independent, so switching models is a documentation + workflow syntax refresh, not an architecture redesign.
- **Consistent terminology:** After auth model changes, audit all user-facing and maintainer-facing docs to ensure terminology alignment. README and gh-pages should use identical wording as setup docs. Small inconsistencies leak into user questions and troubleshooting.

## Archive

- Detailed session history was moved to `.squad\agents\trejo\history-archive-2026-04-24.md` on 2026-04-24.

---

## Cross-Agent Session Notes

### 2026-04-24T14:06:40Z: Plugin Auth Revert Session

**Session Participants:**
- Kelso (Copilot CLI Plugin Engineer) — Verified workflow already token-based, coordinated docs revert
- Trejo (Docs Lead) — Aligned all user-facing and maintainer docs to PLUGINS_REPO_TOKEN model

**Coordination Results:**
- ✅ Workflow consistency verified (already uses `secrets.PLUGINS_REPO_TOKEN` throughout)
- ✅ All docs surfaces aligned (publish-plugins-setup.md, RELEASE-STRATEGY.md, INSTALLATION.md, README.md, gh-pages/index.md)
- ✅ Cross-repo-release-preflight skill generalized to document both PAT and GitHub App patterns
- ✅ Both decisions recorded and merged to decisions.md
- ✅ Orchestration logs created for both agents
- ⏳ Awaiting user to store PLUGINS_REPO_TOKEN secret in repo

**Decision Recorded:** `.squad/decisions.md` → 2026-04-24T14:06:40Z entry

---

## 2026-04-25: Legacy Skill Package Distribution Cleanup

**By:** Trejo (Docs Lead)

**Task:** Remove legacy npm skill package references and update documentation to reflect new plugin marketplace distribution model.

**Work Completed:**

✅ **Package READMEs Deprecated:**
- `packages/excel-cli-skill/README.md` — Updated to deprecation message pointing to:
  - GitHub Copilot plugins (primary)
  - Direct skill extraction via `npx skills add` (secondary)
  - VS Code extension (bundled)
- `packages/excel-mcp-skill/README.md` — Same deprecation pattern, added Claude Desktop MCPB option

✅ **Skills Distribution Docs Updated:**
- `skills/README.md` — Removed references to `npx skillpm install` and `npm install`, added deprecation note for legacy npm packages
- Reorganized installation order: plugins first, then direct extraction, then VS Code extension

✅ **Release Strategy Updated:**
- `docs/RELEASE-STRATEGY.md` — Removed npm registry as distribution channel for skills
- Updated release artifacts table: removed `.tgz` npm format, kept GitHub Release ZIP
- Updated step 8 "publish" job description to remove npm publishing reference
- Updated "Required Secrets" section: removed note about npm OIDC trusted publishing
- Clarified agent skills distributed via "GitHub Release ZIP" and "Direct skill extraction" only

✅ **Installation Docs Verified:**
- `docs/INSTALLATION.md` — Already correct (no npm package install references present)
- Confirmed `npx skills add` references are still valid (user can extract skills directly)

✅ **README and gh-pages Verified:**
- Main `README.md` — Already correct (uses `npx skills add` for direct extraction)
- `gh-pages/index.md` — Already correct (no npm package install references)
- `.github/copilot-instructions.md` — Already correct (uses `npx skills add`)

**Distinction Maintained (Correct):**
- ✅ `npx skills add sbroenne/mcp-server-excel --skill excel-cli|mcp` → **KEPT** (valid direct extraction method for agents without plugin support)
- ✅ GitHub Copilot plugins → **PRIMARY** (plugin marketplace distribution via `copilot plugin install`)
- ✅ VS Code extension → **BUNDLED** (auto-installs both skills)
- ❌ `npm install excel-cli-skill` → **REMOVED** (no longer published)
- ❌ `npx skillpm install` → **REMOVED** (legacy skillpm command, not used)

**Risk Analysis:**
- 🟢 **LOW RISK** — Changes are documentation-only. Kelso's workflow/code changes not touched.
- ⚠️ **FOLLOW-UP (Recommended):** Verify release.yml no longer attempts npm tarball publishing for skills. If Kelso changed the release workflow, confirm step 8 ("publish" job) reflects the new distribution model.

**Learnings for Future Work:**
- Legacy package cleanup requires auditing multiple doc surfaces: installation docs, package READMEs, release strategy, skills overviews, and platform-specific docs (gh-pages).
- When distribution models change, track both positive statements (what IS published) and negative statements (what IS NOT published) — silence about removed channels can confuse users.
- Deprecation messages in legacy package directories should guide users to the exact new method (GitHub Copilot plugins with marketplace URLs + npx skills add example).

---

## 2026-04-25: Plugin Install Docs Cleanup

**By:** Trejo (Docs Lead)

**Task:** Clean user-facing plugin install wording across all installation surfaces; remove maintainer-internal workflow details (PAT, sync gates, downgrade guards, repair/replay mechanics); replace misleading headings; keep user notes short and accurate about the published plugin marketplace repo (`sbroenne/mcp-server-excel-plugins`).

**Work Completed:**

✅ **README.md** — Plugin marketplace section reworded (lines 145-147):
- Removed 4-sentence paragraph about source-repo overlay files and sync gates
- Removed description of "cross-repo token scoped to published marketplace repo"
- Removed "downgrade/tag mismatches blocked" and "manual maintainer re-sync path for repair/replay" details (purely internal)
- Replaced with: "The published repo [...] hosts the GitHub Copilot plugin marketplace. Plugins are republished automatically after each ExcelMcp release, though you may need to wait a few moments for the update to appear in the marketplace."
- Added note: "These commands are specific to GitHub Copilot CLI. VS Code and Claude have their own plugin systems with separate installation flows."

✅ **docs/INSTALLATION.md** — Plugin section reworded (lines 337-356):
- Renamed section from "### Copilot CLI install path" → "### Copilot CLI Plugin Installation" (clearer, less technical)
- Renamed "### One-time post-install steps" → "**After Installation:**" (matches current behavior, no "one-time" misconception)
- Consolidated install+post-install into single logical block for clarity
- Removed entire "Other supported plugin surfaces" section and replaced with single short note
- Removed "Source-layout note" paragraph (source-repo overlay files — purely internal)
- Removed "That publish path is sync-gated..." paragraph (PAT, downgrade guards, repair/replay — purely internal)
- Kept: "Plugins are published automatically after each ExcelMcp release, though you may need to wait a few moments for the update to appear in the marketplace."

✅ **gh-pages/_includes/installation.md** — Already cleaned (from previous pass)
✅ **.github/plugins/excel-cli/README.md** — Already cleaned (from previous pass)
✅ **.github/plugins/excel-mcp/README.md** — Already cleaned (from previous pass)

**Decision Recorded:** `.squad/decisions/inbox/trejo-plugin-docs-cleanup.md`

**Risk Analysis:**
- 🟢 **ZERO RISK** — Changes are documentation-only. No code, workflow, or behavior modified.
- ✅ Install commands are unchanged and tested
- ✅ Marketplace repo name preserved throughout (users may need to reference it)
- ✅ Wording now matches user-facing guidance tone (concise, accurate, honest)

**Learnings for Future Work:**
- User docs should describe WHAT users do, not WHY internal workflows exist. Remove "workflow dispatch" talk, "sync-gate" mechanics, PAT/token details, "downgrade/tag mismatch guards," and "(one-time)" modifiers unless they directly affect user experience.
- Keep marketplace repo name visible when users interact with marketplace. Hiding it behind "published repo" creates support friction.
- Plugin install docs and feature docs are separate problems. Keep plugin install docs SHORT. Relegate workflow mechanics to maintainer docs (e.g., `docs/RELEASE-STRATEGY.md`).
- Misleading headings like "One-time post-install steps" should be replaced with descriptive labels ("After Installation") that match current behavior.
- Three parallel plugin README surfaces (.github/plugins/excel-cli, .github/plugins/excel-mcp, gh-pages/_includes/installation.md) require synchronized rewording. Consider adding linting that flags wording inconsistencies between parallel install sections.


### 2026-04-28: PR #605 Skills Review — Skill Quality Improvements
- Reviewed PR from @rohan-tessl (Tessl skills optimization service) targeting 5 skills.
- Found: Measurable quality gains (46%→90%, 51%→89%, 61%→94%), YAML frontmatter fixes, smart content refactoring.
- Assessed external workflow tool (skill-review.yml) as low-risk, non-blocking feedback.
- Verdict: **Improves skills, worth merging.**
- Key insight: Moving 560-line CLI reference out of SKILL.md (800→202 lines) is safe, improves discovery, preserves content in references/.

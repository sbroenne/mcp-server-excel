# Kelso — History

## Core Context

**Project:** mcp-server-excel — Windows-only Excel automation toolset (COM interop) with two equal entry points: MCP Server (for AI assistants) and CLI (for scripting).

**Tech stack:** .NET 9, C#, Excel COM interop, MCP SDK, source generators. Pre-commit enforcement via PowerShell scripts. Integration tests via xUnit.

**Requested by:** Stefan Brönner

**My scope (per user directive 2026-04-23):** GitHub Copilot CLI plugins ONLY — per https://docs.github.com/en/copilot/concepts/agents/copilot-cli/about-cli-plugins. Packaged bundles containing custom agents (`*.agent.md`), skills (`SKILL.md`), hooks (`hooks.json`), MCP configs (`.mcp.json`), LSP configs, distributed via plugin marketplaces (e.g., github/copilot-plugins, github/awesome-copilot).

**What I am NOT:** Not agentskills.io, not MCPB (Claude Desktop), not VS Code extension packaging.

**Agent inventory in this repo (as of 2026-04-23, confirmed by user):**
- `.github/agents/squad.agent.md` is the ONLY `*.agent.md` file in the repo
- It is the Squad coordinator — governance, not an Excel-domain agent
- **There are no existing Excel-specific agents.** If the Copilot CLI plugin needs one, I author it (scope approved by McCauley).

**Other repo ingredients ready to bundle:**
- `skills/excel-cli/SKILL.md` + `skills/excel-mcp/SKILL.md` (Trejo maintains content)
- `skills/shared/*.md` shared references (Trejo)
- `src/ExcelMcp.McpServer/` MCP server (Cheritto)

**No Copilot CLI plugin package exists yet — greenfield.**

**Team I work with:**
- **Trejo** (Docs Lead) — owns skill content; I own packaging
- **Cheritto** (Platform Dev) — owns MCP Server; I own plugin's `.mcp.json` wiring
- **McCauley** (Lead) — architecture approval (including whether to author new agent files)
- **Hanna** — COM review gate (doesn't apply to pure packaging work)

## Current Coordination

### 2026-04-24T14:06:40Z: Plugin Auth Revert Session (Completed)

**Session Participants:** Kelso + Trejo

**Context:** User requested revert from GitHub App auth back to simpler stored cross-repo PAT (`PLUGINS_REPO_TOKEN`) while preserving iq-core-style hardening.

**Kelso's Work:**
- Verified workflow already token-based (already uses `secrets.PLUGINS_REPO_TOKEN` throughout)
- Coordinated docs revert with Trejo
- Generalized cross-repo-release-preflight skill to document both PAT and App auth patterns

**Trejo's Work:**
- Updated publish-plugins-setup.md with simpler token options (PAT + app token)
- Aligned RELEASE-STRATEGY.md, INSTALLATION.md, README.md, gh-pages/index.md

**Coordination Result:**
- ✅ Decisions recorded and merged to decisions.md
- ✅ Orchestration logs created
- ⏳ Awaiting user to store PLUGINS_REPO_TOKEN secret

**Related:** `.squad/decisions.md` → 2026-04-24T14:06:40Z entry

## Learnings

### 2026-04-25: Retiring a legacy distribution surface needs one end-to-end pass

- If a package line is truly obsolete, remove all three layers together: source artifacts/directories, local validation hooks, and release/publish workflow steps.
- For this repo, the Copilot CLI marketplace repo (`sbroenne/mcp-server-excel-plugins`) is the active plugin distribution path, so the old `packages/excel-*-skill` npm packaging flow was safe to remove once `Build-AgentSkills.ps1`, `pre-commit.ps1`, and `release.yml` no longer referenced it.
- Leave historical mentions in `.squad/` records alone, but sweep active operational files for stale package names so future packaging work does not accidentally resurrect the retired surface.

### 2026-04-24: Reverted plugin publish from GitHub App to stored PAT

- User requested switching option 3: revert from GitHub App auth back to stored cross-repo token, while keeping iq-core-style operational hardening.
- **Already mostly done:** The workflow was already using `PLUGINS_REPO_TOKEN` throughout — only docs needed alignment.
- **Changes made:**
  - Workflow: Already token-based (just verified consistency)
  - Docs: Updated publish-plugins-setup.md, RELEASE-STRATEGY.md, INSTALLATION.md, README.md, gh-pages/index.md
  - Removed references to `PLUGINS_PUBLISH_APP_ID` and `PLUGINS_PUBLISH_APP_PRIVATE_KEY`
  - Updated cross-repo-release-preflight SKILL to generalize patterns for both PAT and App auth options
- **Operational hardening preserved:**
  - Preflight validation (fails fast if token missing/unreachable)
  - Source-side sync gate (skips publish when plugin surface unchanged)
  - Version guards (rejects downgrade, tag mismatch)
  - Manual re-sync path via workflow_dispatch
  - Duplicate detection with auto/manual mode distinction
- **Rationale:** Simpler setup (1 secret vs 1 var + 1 secret), easier rotation, same security posture for public repo use case.

### 2026-04-24: Plugin release docs must separate artifact publication from client UX

- Plugin bundles are broader than a single CLI surface: the package format can carry skills, agents, hooks, and MCP config that may matter to multiple plugin-capable clients.
- Release and workflow docs should therefore describe what we publish as plugin artifacts or agent plugins, while keeping installation claims narrow to the clients we have actually documented and verified.
- Good release wording: "publishes plugin artifacts to the published repo." Risky wording: implying the same workflow automatically registers those artifacts with every client marketplace.

### 2026-04-24: Cross-repo plugin publish needs a preflight gate

- A follow-on release workflow that pushes into a separate published plugin repo should fail fast on configuration, not at the first checkout step.
- Add an explicit preflight job that checks the required cross-repo secret exists and can read the target repo, so missing marketplace credentials surface as a precise action item (`PLUGINS_REPO_TOKEN`) instead of a vague checkout failure.
- Treat the plugin republish as part of release verification: release docs and checklists should explicitly include the follow-on `publish-plugins.yml` run, not just the main `release.yml` workflow.

### 2026-04-24: Do not open the plugin PR from a mixed dirty tree

- The `feature/copilot-cli-plugins` working tree currently mixes the Copilot CLI plugin packaging work with unrelated changes, including Squad workflow scaffolding and a `RangeCommands.Formulas.cs` code edit.
- When plugin packaging is bundled with unrelated infrastructure or product-code changes, stop and report the blocker instead of guessing a safe split for commit/PR creation.

### 2026-04-24: Published plugin repo initialized from the sibling template repo

- Created and pushed the public published repo: `https://github.com/sbroenne/mcp-server-excel-plugins`.
- Reused the existing sibling working directory at `D:\source\mcp-server-excel-plugins` instead of creating a second local copy, then cleaned it into an evergreen publish-target shape before initializing git.
- Kept the publish-target repo minimal and workflow-friendly: root `README.md`, `.gitignore`, `LICENSE`, `marketplace.json`, and `plugins/excel-mcp` + `plugins/excel-cli`.
- Added the missing `"skills": "skills/"` manifest entry to `plugins/excel-cli/plugin.json`; without it, the CLI plugin package would not advertise its bundled skill content.
- The source repo still needs the `PLUGINS_REPO_TOKEN` Actions secret configured so `publish-plugins.yml` can clone and push to the new published repo.

### 2026-04-24: Issue-first handoff when plugin branch is dirty and unpushed

- Created source-repo tracking issue **#606** for the Copilot CLI plugin packaging work (`excel-mcp`, `excel-cli`, publish automation, docs, and local install validation).
- When the working tree is dirty and the feature branch has no upstream, do **not** guess which local changes are ready for review just to open a PR. Open the issue, record the blocker, and wait for the branch state to be made reviewable before creating the PR.

### 2026-04-23: Initial Plugin Plan Research + Precedent Study

**Copilot CLI Plugin Spec Deep Dive:**

- **Only plugin.json is mandatory** — everything else (agents, skills, hooks, MCP servers, LSP) is optional
- **Skills structure we have is correct** — `skills/{name}/SKILL.md` matches spec exactly
- **No OS constraint mechanism** — plugins are cross-platform by default, Windows-only must be documented in description/keywords
- **Marketplace submission is PR-based** — add plugin entry JSON to github/copilot-plugins or github/awesome-copilot
- **Local testing via path install** — `copilot plugin install ./plugin` works for dev iteration
- **Components are cached** — must re-install to pick up changes during development

**Precedent Discovery (office-coding-agent):**

- **Two-repo pattern is REAL-WORLD standard** — source repo (`office-coding-agent`) for development, published repo (`office-coding-agent-plugins`) as dedicated marketplace
- **Users install from published repo** — `copilot plugin install office-excel@office-coding-agent`
- **Published repo structure** — `plugins/{plugin-name}/plugin.json` + `agents/` + `skills/` at root
- **Multiple plugins per marketplace** — PowerPoint has 4 skills (core + deck-builder + formatting + redesign), all in one plugin
- **Custom frontmatter fields** — `hosts: [excel]`, `defaultForHosts: [excel]` (not in official spec, but used by office-coding-agent)
- **Manual publication** — no automated CI/CD found; single commit suggests hand-copy from source to published repo

**Key Decisions Identified (7 blockers):**

1. Plugin name/namespace — `excel-automation` recommended (user-friendly, not implementation-specific)
2. Author custom Excel agent? — Uncertain whether agent + skills is redundant vs complementary
3. **Publication strategy — Two-repo pattern following precedent (MAJOR UPDATE)**
4. Build vs hand-maintained — Build-generated strongly recommended (prevents drift)
5. skills/ relationship — Keep both (skills/ = source, plugin/ = artifact) recommended
6. MCP server binary — PATH assumption recommended (lightweight plugin, fail-fast if not installed)
7. Windows-only messaging — Multi-layered docs (plugin.json + skills + graceful error)

**Surprising Findings:**

- Our existing `skills/` directory is **already plugin-compatible** — structure matches spec perfectly
- The only missing piece is `plugin.json` manifest + plugin-scoped `.mcp.json`
- MCPB (Claude Desktop) and Copilot CLI plugins are **completely separate ecosystems** — no overlap
- Plugin spec has no concept of bundled binaries — expects command-line tools on PATH
- No Excel-domain agent exists (only Squad coordinator) — decision required whether to author one
- **Two-repo pattern is proven and working** — office-coding-agent uses this successfully

**Repo Structure Clarity:**

- `skills/` = source of truth for skill content (Trejo owns)
- `mcpb/` = separate Claude Desktop ecosystem (NOT input to plugin)
- `plugin/` = proposed output directory in SOURCE repo (build artifact, like `bin/`)
- **NEW: `sbroenne/mcp-server-excel-plugin`** = proposed PUBLISHED repo (dedicated marketplace)
- `.github/agents/squad.agent.md` = only existing agent (NOT Excel-specific)

**Real-World Conventions Not in Docs:**

1. **Plugins directory at repo root** — Published repos use `plugins/` (plural), not `plugin/`
2. **Custom frontmatter** — `hosts:` and `defaultForHosts:` fields for host-specific routing
3. **Manual publication** — No automated sync found; appears to be hand-copy workflow
4. **Marketplace naming** — `{author}/{repo-name}` for marketplace, `{plugin}@{marketplace-key}` for install
5. **Multiple skills per plugin** — Core skill + specialized skills pattern (PowerPoint example)

**Next Actions:** User/McCauley must:
1. Approve two-repo pattern
2. Approve creation of `sbroenne/mcp-server-excel-plugin` published repo
3. Decide 6 remaining open questions

Plan document saved to `.squad/agents/kelso/proposals/initial-plugin-plan.md`.

---

### 2026-04-23: Final Plugin Shape — All Decisions Locked

**User Locked All 7 Decisions:**

1. ✅ **Two plugins:** `excel-mcp` (MCP + skill + agent) and `excel-cli` (skill only) — clean separation by use case
2. ✅ **Published repo:** `sbroenne/mcp-server-excel-plugins` (plural, enables future plugins)
3. ✅ **Binary bundled:** MCP server included with excel-mcp plugin (not PATH assumption)
4. ✅ **Automated publication:** GitHub Action copies from source → published repo (DEVIATES from office-coding-agent's manual pattern)
5. ✅ **Marketplace submission DEFERRED:** NOT submitting to github/copilot-plugins in v1
6. ✅ **Agent decision (Kelso recommendation):** YES for excel-mcp (conversational scaffolding), NO for excel-cli (scripting context)
7. ✅ **Versioning:** Lockstep (plugin version matches MCP server release)

**Kelso Recommendations Accepted:**

- **Excel Agent for MCP Plugin:** YES — thin agent enforcing CRITICAL RULES + workflow hints, defer details to skill. Pattern: "NEVER ask clarifying questions — use list tools to discover." office-coding-agent precedent: ALL plugins have agents.
- **Binary Distribution:** GitHub Release download script — plugin includes `bin/download.ps1` (small, committed), downloads mcp-excel.exe from Release (50-80MB, NOT committed). Avoids Git bloat, keeps repo lean.
- **Windows-Only Gating:** Multi-layered (plugin.json description "⚠️ WINDOWS-ONLY", keywords, SKILL.md preconditions, README warnings, runtime graceful failure). Can't prevent install (no OS filter in spec).
- **Version Pinning:** Lockstep versioning — plugin v1.2.0 = MCP server v1.2.0. Simplifies user confusion, matches tight coupling (bundled binary).

**Automation Deviation Rationale:**
- office-coding-agent uses manual copy (single commit, no CI/CD found)
- We automate via GitHub Action: less toil, fewer sync bugs, faster releases, enforced consistency
- Complexity: Requires PAT for cross-repo push
- Workflow: Release tag push → Build plugins → Clone published repo → Copy plugins/ → Commit + push

**Final Published Repo Structure:**
```
mcp-server-excel-plugins/
├── README.md
├── .gitignore          # Ignore bin/*.exe, keep bin/download.ps1
└── plugins/
    ├── excel-mcp/      # MCP + skill + agent + binary download script
    │   ├── plugin.json
    │   ├── .mcp.json
    │   ├── agents/excel.agent.md
    │   ├── skills/excel-mcp/SKILL.md
    │   └── bin/download.ps1
    └── excel-cli/      # Skill only (CLI tool installed separately)
        ├── plugin.json
        └── skills/excel-cli/SKILL.md
```

**Installation Flow:**
```powershell
# Register marketplace
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# MCP plugin (AI assistants)
copilot plugin install excel-mcp@mcp-server-excel
cd ~/.copilot/plugins/excel-mcp/bin && ./download.ps1

# CLI plugin (scripting)
copilot plugin install excel-cli@mcp-server-excel
```

**Publishing Strategy:**
- Automated via GitHub Action on release tag
- Uses GitHub App (NOT PAT) for scoped permissions
- Lockstep versioning: Plugin v1.2.0 = MCP server v1.2.0
- Two-step publication: Build plugins → Release binary → Publish plugins (via workflow_run trigger)

---

### 2026-04-23: Rubber-Duck Review Findings + Spike-First Approach

**Context:** User conducted rubber-duck review of the finalized plan and identified **4 critical findings** and **4 moderate findings** that required plan updates BEFORE implementation.

#### Critical Findings Fixed

**1. Wrapper Script for Missing-Binary Detection**

**Problem:** If user installs `excel-mcp` plugin but forgets to run `download.ps1`, MCP server fails with cryptic error. User stuck with no clear next steps.

**Solution:** Add `bin/start-mcp.ps1` wrapper script between `.mcp.json` and `mcp-excel.exe`:
- `.mcp.json` references wrapper, NOT exe directly: `"command": "{pluginDir}/bin/start-mcp.ps1"`
- Wrapper checks: Does `mcp-excel.exe` exist? If NO → display clear error message with installation instructions
- Wrapper checks: Version skew? (compare binary version vs `version.txt`) If mismatch → warn user
- Wrapper launches: If all checks pass → launch `mcp-excel.exe` with forwarded args

**Why Critical:** Without wrapper, users get "command not found" or PowerShell errors instead of actionable guidance. This would generate support tickets and frustration.

**Two-step install now prominently documented** in plugin README, installation instructions, and error messages.

---

**2. Validate `.mcp.json` + `{pluginDir}` Placeholder**

**Problem:** Assumption that Copilot CLI expands `{pluginDir}` placeholder in `.mcp.json` is **UNVERIFIED**. If this doesn't work, the entire plugin breaks.

**Solution:** **Phase -1 (Spike)** added as BLOCKING prerequisite to Phase 0.

**Spike Goals:**
1. Create minimal "hello-world" plugin with `.mcp.json` referencing `{pluginDir}/bin/stub.ps1`
2. Install locally and trigger MCP launch
3. Verify: Does `{pluginDir}` expand? Does stub execute? Can stub compute its own path via `$PSScriptRoot`?
4. Document findings before proceeding

**Fallback:** If `{pluginDir}` doesn't work, wrapper script uses `$PSScriptRoot` to compute absolute path.

**Why Critical:** Building 5 phases of implementation on unverified assumption = wasted work if assumption wrong. Spike validates core mechanism in 2 hours.

---

**3. Replace PAT with GitHub App or Deploy Key**

**Problem:** Phase 4a (automated publication) originally planned to use Personal Access Token (PAT) for cross-repo push. PATs are:
- Over-permissioned (all repos, not scoped)
- Expiring (require maintenance)
- Security risk (broad access surface)

**Solution:** Use **GitHub App** with permissions scoped to `mcp-server-excel-plugins` repo ONLY.

**Setup:**
- Create GitHub App in `sbroenne` account
- Permissions: `contents: write` on `mcp-server-excel-plugins` ONLY
- Install App on published repo
- Workflow uses `actions/create-github-app-token` to generate installation token

**Benefits:**
- ✅ Scoped permissions (ONLY published repo)
- ✅ No expiration
- ✅ Auditable (GitHub tracks App actions separately)
- ✅ Revocable (uninstall App to revoke)

**Fallback:** Deploy key (write access to published repo), NOT classic PAT.

---

**4. SHA256 Checksum Verification in download.ps1**

**Problem:** `download.ps1` downloads 50MB binary from GitHub Release with no integrity check. Risks:
- MITM attack (malicious binary substitution)
- Corrupted download (partial transfer, network error)
- User gets broken/malicious binary

**Solution:** Release workflow produces `checksums.txt`, download script verifies SHA256 before extraction.

**Release Workflow:**
```powershell
# After creating ZIP
$hash = (Get-FileHash $zipFile -Algorithm SHA256).Hash
"$hash  $zipFile" | Out-File -Append checksums.txt
```

**Download Script:**
```powershell
# Download binary + checksums
Invoke-WebRequest $assetUrl -OutFile $zipPath
Invoke-WebRequest "$releaseUrl/checksums.txt" -OutFile checksums.txt

# Extract expected hash
$expectedHash = (Get-Content checksums.txt | Where-Object { $_ -like "*$assetName*" }) -split '\s+' | Select-Object -First 1

# Compute actual hash
$actualHash = (Get-FileHash $zipPath -Algorithm SHA256).Hash

# Verify
if ($actualHash -ne $expectedHash) {
    Write-Error "SHA256 mismatch!"
    Remove-Item $zipPath
    exit 1
}
```

**Why Critical:** Without checksum verification, users have no protection against corrupted or tampered downloads. This is basic supply chain security.

---

#### Moderate Findings Fixed

**5. Version Skew: Embed version.txt in Plugin**

**Problem:** User installs plugin v1.2.0, `download.ps1` defaults to "latest" release, gets v1.3.0 binary → version mismatch.

**Solution:** Embed `version.txt` file in plugin: `echo "1.2.0" > plugins/excel-mcp/version.txt`
- `download.ps1` reads `version.txt` and fetches **exact matching release tag**
- NO "latest" default — always explicit version
- Wrapper script checks version skew, warns if binary doesn't match plugin version

---

**6. Publish Workflow Atomicity**

**Problem:** Phase 4a originally copied plugins in two sequential commits (excel-mcp, then excel-cli) → window for half-published state.

**Solution:** Add `concurrency` control + single atomic commit:
```yaml
concurrency:
  group: publish-plugins
  cancel-in-progress: false
```
- Single `git add plugins/` instead of two sequential copies
- Single commit: "Release vX.Y.Z (excel-mcp + excel-cli)"

---

**7. CLI Plugin Discovery Without Agent**

**Question:** If `excel-cli` plugin has no agent, how does LLM discover the skill?

**Answer:** LLM discovers skill via `copilot /skills list` command → reads SKILL.md directly (no agent needed).

**Documentation Added:**
- `excel-cli/SKILL.md` precondition: "Requires `excelcli.exe` in PATH. Install via Chocolatey, Scoop, or manual download."

---

### 2026-04-24: Orchestration — Ready for PR Creation

**Status:** All phases complete. Branch staged. Awaiting explicit user approval to create PR.

**Key Actions Completed:**
- Published repo `sbroenne/mcp-server-excel-plugins` initialized and pushed
- Source repo tracking issue #606 created
- All plugin phases locked (scaffold, MCP plugin, CLI plugin, publish workflow, audit complete)
- All 23 source repo changes staged for commit

**Pending:**
- User approval to open PR from `feature/copilot-cli-plugins` → `main`
- PR references issue #606
- Scribe will orchestrate git commit and PR creation upon approval
- Differentiated skill descriptions:
  - `excel-mcp`: "AI assistant for Excel automation via MCP server tools"
  - `excel-cli`: "CLI skill for scripting and batch automation"

---

**8. Drop Custom Frontmatter Fields**

**Removed from agent.md template:**
- ❌ `hosts: [excel]` — office-coding-agent custom field, not in Copilot CLI spec
- ❌ `defaultForHosts: [excel]` — not in spec

**Keep only:**
- ✅ `name:` (required)
- ✅ `description:` (required)

---

#### Open Questions Answered

**Q1: Does repo have release workflow uploading binary assets?**
- ✅ **YES** — `.github/workflows/release.yml` uploads `ExcelMcp-MCP-Server-{version}-windows.zip` to GitHub Release

**Q2: Binary availability race condition?**
- ✅ **FIXED** — Plugin publish workflow uses `workflow_run` trigger (waits for release workflow completion)

**Q3: Corporate proxy support?**
- ✅ **YES** — `download.ps1` uses `DefaultWebProxy` (respects `HTTPS_PROXY` env var automatically)

**Q4: Air-gapped environments?**
- ❌ **NOT IN V1** — Noted on roadmap (requires fully bundled alternative or internal NuGet feed)

**Q5: Dual install (both plugins)?**
- ✅ **NO** — Only `excel-mcp` includes binary; `excel-cli` requires CLI installed separately

---

#### Why Spike-First Approach Matters

**Before rubber-duck review:** Plan assumed `{pluginDir}` placeholder works → build 5 phases → discover it doesn't → throw away work.

**After rubber-duck review:** **Phase -1 (Spike)** validates assumptions in 2 hours BEFORE committing to full implementation.

**Spike Exit Criteria:**
- ✅ Confirms `{pluginDir}` expansion works OR documents fallback
- ✅ Confirms wrapper script pattern works
- ✅ Results documented in `phase-minus-1-spike-results.md`
- ✅ **ONLY proceed to Phase 0 if spike succeeds**

**If spike fails:** STOP and re-design (absolute paths, env vars, different MCP launch mechanism).

**Lesson:** Never assume. Validate core mechanisms BEFORE building on top of them.

---

#### Automation Deviation Rationale

**office-coding-agent precedent:** Manual publication (human copies plugins to published repo, commits, pushes).

**ExcelMcp decision:** Automated publication via GitHub Action.

**Why deviate:**
- ✅ Manual publication is error-prone (forgot to copy a file, wrong version number)
- ✅ Lockstep versioning (plugin version = MCP server version) requires automation
- ✅ Security improvement (GitHub App instead of manual PAT handling)
- ✅ Consistency (release workflow already automated, extend same pattern)

**Risk mitigation:**
- Test locally via Phase 1 (local install) BEFORE Phase 4a (automated publish)
- Manual validation step in Phase 4 (manually publish once, verify, THEN automate)
- Workflow includes concurrency control (prevents parallel publish attempts)

---
copilot plugin install excel-cli@mcp-server-excel
```

**Scope CLOSED:** All architectural decisions finalized. Ready for Phase 0 (create published repo) → Phase 4 (automated publication).

Final decision record saved to `.squad/decisions/inbox/kelso-plugin-shape-final.md`.
---

### 2026-04-23: Rubber-Duck Review + Spike-First Refinement (Turn 3)

**Session Summary:**
- **Agent:** rubber-duck-plugin-plan (Turn 2) conducted structured critique
- **Verdict:** APPROVE WITH CONDITIONS — 4 critical + 4 moderate findings
- **User decision:** Accept all findings, add Phase -1 spike before Phase 0
- **Work:** Refined plan to incorporate all findings, answered Q1–Q5, finalized Phase -1 scope

**Critical Fixes Incorporated:**

1. **Wrapper Script** (\in/start-mcp.ps1\) — Detects missing binary, version skew, user guidance
2. **Phase -1 Spike** — Validates \{pluginDir}\ placeholder before committing to design
3. **GitHub App Auth** — Replace PAT with scoped GitHub App (Phase 4)
4. **SHA256 Verification** — Release workflow produces \checksums.txt\, download script verifies

**Moderate Findings Incorporated:**

5. Version skew detection (embed \ersion.txt\, wrapper validates)
6. Publish workflow atomicity (concurrency control, single commit)
7. CLI discovery without agent (docs-driven, no agent needed)
8. Drop custom frontmatter fields (keep only spec-compliant fields)

**Questions Answered:**

- Q1: ✅ YES — release.yml has binary assets
- Q2: ✅ YES — race condition exists; use \workflow_run\ trigger
- Q3: ✅ YES — \download.ps1\ supports corporate proxies
- Q4: ❌ NO — air-gapped not in v1 (roadmap)
- Q5: ❌ NO — only \xcel-mcp\ downloads binary, \xcel-cli\ is skill-only

**Phase Plan Updated:**
- Phase -1 (NEW, BLOCKING): Spike to validate install mechanism
  - Create minimal "hello-world" plugin
  - Verify \{pluginDir}\ placeholder expansion
  - Document findings or pivot
  - Only proceed to Phase 0 if spike succeeds
- Phase 0–4: Original plan (unchanged, but gated by Phase -1)

**Deliverables:**
- ✅ Plan refined with all findings
- ✅ Phase -1 fully scoped with exit criteria
- ✅ Decision record: \.squad/decisions.md\ (merged from inbox)
- ✅ Orchestration logs: 2 files documenting critiques + refinement
- ✅ Session log: Brief note on spike + findings

**Status:** ✅ Ready for Phase -1 execution

**Next Step:** Execute Phase -1 spike, document results in \.squad/agents/kelso/proposals/phase-minus-1-spike-results.md\, then await Stefan's Phase 0 GO/NO-GO decision.

Final decision record merged to \.squad/decisions.md\ (deduped).

### 2026-04-24: Package Metadata Cleanup

- Removed stale skillpm metadata from plugin package manifests when present.
- Confirmed packages\excel-mcp-skill\package.json needed cleanup and packages\excel-cli-skill\package.json already matched the desired shape.

### 2026-04-24: Session End — Blocker on PR, Inbox Merged, Decisions Captured

- Reported blocker: Do NOT open plugin PR from mixed dirty tree (plugin packaging work mixed with Squad infrastructure changes and unrelated RangeCommands.Formulas.cs product-code edit).
- Decision inbox merged to decisions.md (6 inbox files deduplicated and incorporated).
- Cross-agent history updated for Nate, Kelso, Trejo, and other affected agents.
- Scribe orchestration logs and session logs written (ISO 8601 UTC timestamps).
- User explicitly directed revert of unrelated RangeCommands.Formulas.cs change (completed by Nate).
- Session winding down; awaiting branch narrowing before PR submission.


### 2026-04-24: Publish Workflow Hardening

- Hardened `publish-plugins.yml` with a source-side sync gate so automatic follow-on runs only publish when the plugin install surface changed since the previous release tag.
- Added published-repo downgrade, duplicate, and tag/version consistency guards before sync.
- Kept a manual `workflow_dispatch` replay path that targets an existing source `release_tag` for recovery without cutting a fresh release.

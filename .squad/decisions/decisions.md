# Decisions — ExcelMcp Squad

## 2026-04-25: CLI Tool Bundled in Excel-CLI Skill — Documentation Alignment

**Date:** 2026-04-25T14:22:00Z  
**Decision Owner:** Trejo (Docs Lead)  
**Status:** Complete  
**Affected Surfaces:** README.md, docs/INSTALLATION.md, docs/RELEASE-STRATEGY.md, skills/excel-cli/

### Problem Statement
The `excel-cli` skill needs to ship the CLI binary (`excelcli.exe`) itself for self-contained distribution. Documentation across README, INSTALLATION, RELEASE-STRATEGY, and skill guides did not clearly explain that the CLI is **bundled with the skill** as the primary distribution method.

### Solution Implemented
Updated all user-facing and maintainer docs to clarify the bundled CLI model:

1. **skills/excel-cli/README.md** — "CLI included with skill" + three clear installation methods (skill-first, manual, NuGet)
2. **skills/excel-cli/SKILL.md** — "CLI included with skill — Install via this skill package (no separate download needed)"
3. **docs/RELEASE-STRATEGY.md** — Explicit statement that CLI ships in three ways: ZIP, NuGet, bundled inside excel-cli skill
4. **docs/INSTALLATION.md** — "Optional: CLI Installation" with skill-first guidance
5. **README.md** — "CLI vs MCP Server" section clarified with bundled CLI note

### Key Messaging
Users understand:
- **Single binary, three distribution channels**
- **Via Skill (Recommended):** npx skills add + ZIP + VS Code Extension + GitHub Copilot plugin
- **Standalone:** ExcelMcp-CLI ZIP from GitHub Releases
- **NuGet:** dotnet tool install (requires .NET 10 runtime)

### Files Modified
1. skills/excel-cli/README.md
2. skills/excel-cli/SKILL.md
3. docs/RELEASE-STRATEGY.md
4. docs/INSTALLATION.md
5. README.md

### Verification
- ✅ Skill docs lead users to bundled CLI first
- ✅ Release docs show CLI bundled in skill packages
- ✅ Installation guide explains CLI included with skill
- ✅ README clarifies CLI bundled in skill distribution
- ✅ Consistent terminology across all surfaces

---

## 2026-04-25: Plugin Marketplace Documentation Audit & Alignment

**Date:** 2026-04-25T10:15:00Z  
**Agent:** Trejo (Docs Lead)  
**Status:** Complete  
**Scope:** User-facing + maintainer docs

### Decision
ExcelMcp plugin marketplace documentation aligned across all surfaces. Two-plugin distribution story (`excel-mcp` + `excel-cli`) clearly documented with consistent terminology and accurate counts.

### What Changed
**User-Facing Docs:**
- **docs/AWESOME-COPILOT-PROPOSAL.md** (Renamed) — "GitHub Copilot Plugin Distribution"; removed outdated awesome-copilot registry process; added two-repo architecture explanation; fixed tool count to 25 tools (230 operations)
- **docs/INSTALLATION.md** — Promoted GitHub Copilot Plugins to first dedicated section; clearly labeled both plugins; explained `excel-cli` plugin includes bundled binary
- **skills/README.md** — Listed GitHub Copilot Plugins as recommended install path; clarified relationship: plugins (distribution) bundle skills (behavioral guidance)

**Maintainer Docs:**
- **.github/workflows/docs/publish-plugins-setup.md** — Updated overview to explain two-repo pattern; added reference to user-facing plugin docs; clarified published repo is authoritative

### Key Messaging
Users understand:
1. **Two plugins** (excel-mcp + excel-cli) in GitHub Copilot plugin marketplace
2. **Excel MCP** — Full MCP Server with 25 tools (230 operations) for conversational AI
3. **Excel CLI** — CLI tool with bundled skill; includes binary (no separate download)
4. **Source → Published** — This repo syncs to `sbroenne/mcp-server-excel-plugins` (published marketplace)
5. **Auto-update** — Both plugins republished automatically after each source release

### Validation Completed
- ✅ Tool/operation counts match FEATURES.md
- ✅ No stale "excel-automation" references
- ✅ Two-repo, two-plugin pattern clear across all surfaces
- ✅ Skills vs plugins distinction clarified
- ✅ GitHub Copilot Plugins documented as primary installation method

---

## 2026-04-24: Bundle the Self-Contained CLI Inside the Excel-CLI Plugin

**Decision Owner:** Kelso (Copilot CLI Plugin Engineer)

### Decision
The `excel-cli` Copilot plugin should publish with the actual self-contained `excelcli.exe` deliverable in `plugins/excel-cli/bin/`, instead of remaining a skill-only wrapper that assumes a separate CLI install.

### Why
- Existing plugin packaging story incomplete: install docs for `excel-cli` better than artifact itself
- Repo already produces self-contained CLI publish output in release flows; reusing during plugin publication lower risk
- Shipping binary inside plugin keeps Copilot plugin install flow honest while preserving standalone ZIP and NuGet paths

### Implementation Shape
1. `publish-plugins.yml` publishes self-contained CLI into staging folder before plugin build
2. `scripts/Build-Plugins.ps1` copies validated templates, applies source-owned overlays from `.github/plugins/`, bundles staged CLI output into `plugins/excel-cli/bin/`
3. Plugin includes one-time `install-global.ps1` helper that writes `excelcli` shims into `~/.copilot/bin` and adds directory to user PATH

### Implications
- Plugin-facing source diffs must include `.github/plugins/**` in publish sync gate
- Docs should describe Copilot plugin path as bundling self-contained CLI; keep broader client-install claims narrow

---

## 2026-04-24: Canonical Marketplace Manifest with Legacy Fallback

**Decision Owner:** Kelso (Copilot CLI Plugin Engineer)

### Context
Official Copilot CLI marketplace docs expect manifest at `.github/plugin/marketplace.json` with `metadata.version` and per-plugin `source` paths. Published plugin repo currently exposes legacy root `marketplace.json` with top-level `version` field. Source repo workflows/docs written around legacy shape.

### Decision
Treat published repo as real marketplace (not source repo). Update source-side automation and docs to:
1. Prefer `.github/plugin/marketplace.json` when it exists
2. Fall back to legacy root `marketplace.json` while published repo still in that shape
3. Clearly document that `.github/plugins/` in source repo contains overlays only, not installable plugin roots

### Why
Smallest change making source repo spec-aligned without breaking current published repo. Removes misleading local-source expectations and gives published repo clean migration path.

### Follow-up
When published repo updated, move its manifest to `.github/plugin/marketplace.json`, switch plugin entries to canonical `source` fields, then remove legacy fallback from `publish-plugins.yml`.

---

## 2026-04-24: Reuse Current GitHub Auth Token for Plugin Publish Secret

**Decision Owner:** Kelso (Copilot CLI Plugin Engineer)

### Context
User requested repo secret configured immediately using current GitHub auth context, verified without exposing token value.

### Decision
Use active `gh` authentication context for account `sbroenne` to populate source-repo Actions secret `PLUGINS_REPO_TOKEN` on `sbroenne/mcp-server-excel`.

### Why
- Active auth context already has required repository access and scopes
- `publish-plugins.yml` already aligned to single-secret model; doesn't need older GitHub App variable/secret pair
- Keeps publish path on simpler stored-token setup user explicitly requested

### Verification
- ✅ PLUGINS_REPO_TOKEN appears in source repo's Actions secret list
- ✅ Active `sbroenne` auth context has `admin` permission on `sbroenne/mcp-server-excel-plugins`
- ✅ Source repo Actions settings compatible with workflow

---

## Archive
See `.squad/decisions/archive/` for decisions older than 30 days.

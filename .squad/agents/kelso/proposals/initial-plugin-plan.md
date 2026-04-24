# Copilot CLI Plugin — Final Plan

**Date:** 2026-04-23 (Updated after rubber-duck review + decision lock)  
**Agent:** Kelso  
**Requested by:** Stefan Brönner

---

## Executive Summary

This plan outlines how to transform ExcelMcp's existing skills and MCP server into **TWO separate** distributable GitHub Copilot CLI plugins (`excel-mcp` and `excel-cli`). The plugins will be published to a dedicated marketplace repo (`sbroenne/mcp-server-excel-plugins`) and installable via `copilot plugin install`.

**Status:** FINALIZED + RUBBER-DUCK FIXES. All architectural decisions locked. **Phase -1 (Spike) must validate install mechanism BEFORE Phase 0.**

**Latest Updates (Post Rubber-Duck Review):**
- ✅ Added **Phase -1 (Spike)** to validate `{pluginDir}` placeholder + wrapper script pattern
- ✅ Critical fix: Wrapper script (`bin/start-mcp.ps1`) for missing-binary detection + version skew
- ✅ Critical fix: SHA256 checksum verification in `download.ps1`
- ✅ Critical fix: Replace PAT with GitHub App or deploy key
- ✅ Moderate fix: Embed `version.txt` in plugin for explicit version pinning
- ✅ Moderate fix: Workflow atomicity (concurrency control)
- ✅ Moderate fix: CLI plugin discovery without agent (documented in SKILL.md)
- ✅ Answered Q1-Q5 (release workflow, binary race, proxy support, air-gapped, dual install)

**Key Decisions (LOCKED):**
- Two separate plugins: `excel-mcp` (MCP server + skill) and `excel-cli` (CLI skill only)
- MCP server binary **bundled** with `excel-mcp` plugin (via GitHub Release download script)
- Automated publication via GitHub Action (trigger on release tag)
- Marketplace submission DEFERRED (not part of v1 scope)

---

## Final Plugin Architecture (LOCKED DECISIONS)

### Two-Repo Pattern + Two Separate Plugins

**Source Repo:** `sbroenne/mcp-server-excel` (this repo)
- Develop plugins in `plugins/` directory (build artifacts, gitignored)
- Build script generates both plugins from sources
- Test locally: `copilot plugin install ./plugins/excel-mcp`

**Published Repo:** `sbroenne/mcp-server-excel-plugins` (NEW, create this)
- Dedicated marketplace repo (mirrors office-coding-agent-plugins pattern)
- Structure:
  ```
  mcp-server-excel-plugins/
  ├── README.md              # Installation instructions + two-step install warning
  └── plugins/
      ├── excel-mcp/         # MCP Server plugin
      │   ├── plugin.json    # name: "excel-mcp"
      │   ├── .mcp.json      # References bin/start-mcp.ps1 (WRAPPER, not mcp-excel.exe directly)
      │   ├── version.txt    # "1.2.0" — explicit version for download.ps1
      │   ├── bin/
      │   │   ├── start-mcp.ps1        # Wrapper: detects missing/mismatched binary, launches MCP
      │   │   ├── download.ps1         # Binary downloader with SHA256 verification
      │   │   └── mcp-excel.exe        # Downloaded from GitHub Release (NOT committed)
      │   ├── agents/
      │   │   └── excel.agent.md       # YES - thin workflow orchestrator
      │   └── skills/
      │       └── excel-mcp/
      │           └── SKILL.md
      └── excel-cli/         # CLI plugin
          ├── plugin.json    # name: "excel-cli"
          └── skills/
              └── excel-cli/
                  └── SKILL.md
  ```

**User Installation:**
```powershell
# Register marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install MCP plugin (for AI assistants) — TWO-STEP install!
copilot plugin install excel-mcp@mcp-server-excel

# Step 2: Download MCP server binary (REQUIRED)
cd ~/.copilot/plugins/excel-mcp/bin
./download.ps1  # Downloads mcp-excel.exe from GitHub Release with SHA256 verification

# OR install CLI plugin (for scripting/automation) — NO binary needed
copilot plugin install excel-cli@mcp-server-excel
# NOTE: Requires excelcli.exe installed separately (Chocolatey, Scoop, or manual)
```

**IMPORTANT: excel-mcp plugin requires TWO-STEP installation** (plugin install + binary download). This is prominently documented in plugin README and error messages.

### Plugin Breakdown

| Plugin | Name | Contains | Target Use Case | Binary? |
|--------|------|----------|-----------------|---------|
| **excel-mcp** | `excel-mcp` | MCP server config + excel-mcp skill + Excel agent | Conversational AI assistants (Claude, Copilot Chat) | ✅ YES (via download script) |
| **excel-cli** | `excel-cli` | excel-cli skill only | Scripting, automation, batch operations | ❌ NO (CLI installed separately) |

**Why Two Plugins:**
- Clean separation of concerns (MCP server vs CLI tool)
- Users install ONLY what they need (AI assistant vs automation script)
- Independent versioning if needed (though lockstep recommended)
- Matches our existing skills separation (excel-mcp vs excel-cli)

**Why NOT Single Plugin:**
- Bundling MCP server binary in CLI plugin makes no sense (CLI doesn't use MCP)
- CLI skill references `excelcli.exe` which users install separately
- Two use cases are distinct (conversation vs scripting)

---

## Critical Design Decisions (LOCKED + RECOMMENDATIONS)

### Decision: Custom Excel Agent for MCP Plugin

**Recommendation:** ✅ **YES** — Include `agents/excel.agent.md` in `excel-mcp` plugin, **NO** agent in `excel-cli` plugin

**Rationale:**

**FOR MCP Plugin:**
- MCP tools have rich schemas, but an agent provides **conversational scaffolding**
- Agent enforces CRITICAL RULES without duplicating skill content (e.g., "NEVER ask clarifying questions — use list tools")
- Workflow orchestration: "For Power Query: start with `connection list`, then `power-query import`"
- Personality/voice: "You are an Excel automation expert using the excel-mcp MCP server tools"
- office-coding-agent precedent: ALL plugins have agents (excel.agent.md, powerpoint.agent.md, etc.)
- Agent should be **thin** — enforce rules + workflow hints, defer operational details to skill

**AGAINST CLI Plugin:**
- CLI plugin is for scripting/automation, not conversation
- No conversational context exists (agent would never be invoked)
- CLI skill already provides complete workflow guidance

**Agent Content Pattern (for excel-mcp):**
```markdown
---
name: Excel
description: AI assistant for Excel automation via MCP server tools
hosts: [excel]  # Optional custom field
---

You are an Excel automation expert using the excel-mcp MCP server tools.

CRITICAL RULES:
1. NEVER ask clarifying questions — use list tools (file list, table list, worksheet list) to discover
2. ALWAYS close sessions (file close with save: true) to avoid locking files
3. For bulk operations, use calculation_mode to disable auto-recalc during writes
4. ALWAYS end with a text summary confirming what was done

For complete workflows, gotchas, and operational patterns, see the excel-mcp skill.
```

### Decision: MCP Server Binary Distribution Strategy

**Problem:** .NET self-contained publish outputs are **50-80MB** (typical for Windows x64). Git-based plugin distribution struggles with large binaries (slow clones, repo bloat).

**Locked Decision:** ✅ **Bundle binary** with plugin (Stefan's directive)

**Implementation Strategy:** **GitHub Release Download Script** (recommended)

**How It Works:**
1. Plugin includes `bin/download.ps1` script (small, committed to Git)
2. `.mcp.json` references `bin/mcp-excel.exe` (NOT committed, in `.gitignore`)
3. On plugin install, Copilot CLI runs `bin/download.ps1` (if supported) OR user runs manually
4. Script pulls `mcp-excel.exe` from latest GitHub Release asset
5. Script places binary in `bin/` directory
6. MCP server starts via local path: `{pluginDir}/bin/mcp-excel.exe`

**Tradeoffs:**

| Approach | Pros | Cons |
|----------|------|------|
| **Release Download (RECOMMENDED)** | ✅ Repo stays lean (no 50MB binary in Git)<br>✅ Fast plugin installs<br>✅ Binaries don't pollute Git history | ⚠️ Two-step install (plugin + binary download)<br>⚠️ Requires manual script execution if Copilot CLI doesn't support post-install hooks |
| **Direct Commit** | ✅ Single-step install<br>✅ No external dependencies | ❌ 50-80MB binary per release in Git<br>❌ Slow clone times<br>❌ Repo bloats over time |

**Why Release Download Wins:**
- Git repos are designed for text, not large binaries
- Each release adds 50-80MB to Git history (non-recoverable)
- Plugin updates would re-download binary from Release anyway
- office-coding-agent doesn't bundle Office.js runtime (separate concern)

**Implementation Details:**
```powershell
# bin/download.ps1
param([string]$Version = "latest")

$owner = "sbroenne"
$repo = "mcp-server-excel"
$assetPattern = "ExcelMcp-MCP-Server-*-windows.zip"

# Fetch latest release info from GitHub API
$release = if ($Version -eq "latest") {
    Invoke-RestMethod "https://api.github.com/repos/$owner/$repo/releases/latest"
} else {
    Invoke-RestMethod "https://api.github.com/repos/$owner/$repo/releases/tags/v$Version"
}

# Find MCP server asset
$asset = $release.assets | Where-Object { $_.name -like $assetPattern } | Select-Object -First 1

if (-not $asset) {
    Write-Error "MCP server binary not found in release $($release.tag_name)"
    exit 1
}

# Download and extract
$zipPath = "$PSScriptRoot\mcp-server.zip"
Invoke-WebRequest $asset.browser_download_url -OutFile $zipPath
Expand-Archive $zipPath -DestinationPath $PSScriptRoot -Force
Remove-Item $zipPath

Write-Host "✅ MCP server $($release.tag_name) installed to $PSScriptRoot"
```

**`.mcp.json` Config:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "{pluginDir}/bin/mcp-excel.exe",
      "args": [],
      "env": {}
    }
  }
}
```

**Plugin README Must Document:**
```markdown
## Installation

1. Register marketplace:
   ```powershell
   copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
   ```

2. Install plugin:
   ```powershell
   copilot plugin install excel-mcp@mcp-server-excel
   ```

3. Download MCP server binary (one-time):
   ```powershell
   cd ~/.copilot/plugins/excel-mcp/bin
   ./download.ps1
   ```

**Windows + Excel 2016+ required.**
```

### Decision: Windows-Only Gating Strategy

**Problem:** Excel COM automation is Windows-only. Copilot CLI plugins have no OS constraint mechanism. macOS/Linux users may attempt to install.

**Multi-Layered Approach:**

**1. Pre-Install (Documentation):**
```json
// plugin.json
{
  "name": "excel-mcp",
  "description": "⚠️ WINDOWS-ONLY: Excel automation via COM interop (requires Excel 2016+)",
  "keywords": ["excel", "windows", "windows-only", "com-interop", "automation"]
}
```

**2. During Install (README):**
```markdown
# excel-mcp Plugin

⚠️ **WINDOWS-ONLY** — Requires Windows OS + Microsoft Excel 2016+

This plugin uses Excel COM automation and **will not work** on macOS or Linux.
```

**3. At Runtime (Graceful Failure):**
- MCP server startup: Check for COM availability
- If COM not available: Return clear error via MCP error channel
  ```json
  {
    "error": {
      "code": -32000,
      "message": "Excel COM not available. This MCP server requires Windows + Excel 2016+."
    }
  }
  ```

**4. Skill Preconditions:**
```markdown
## Preconditions

- ⚠️ **Windows OS** — macOS and Linux are NOT supported
- Microsoft Excel 2016 or later installed
- Excel COM interop available (desktop Excel, not Office Online)
```

**Why This Works:**
- Can't prevent install (no OS filter in plugin spec)
- Can gracefully no-op with clear errors
- Documentation at multiple touch points (discovery → install → use)
- Users know BEFORE wasting time

### Decision: Version Pinning Strategy

**Locked Decision:** **Lockstep versioning** — plugin version matches MCP server release version

**Why:**
- Simplifies user confusion ("I have plugin v1.2.0, what server version?" → "Same version")
- Plugin is tightly coupled to MCP server (binary bundled, .mcp.json references it)
- Each plugin release corresponds to one MCP server release
- Precedent: office-coding-agent likely uses lockstep (tight coupling between add-in and plugins)

**Implementation:**
- Plugin `plugin.json` version: `"version": "1.2.0"`
- Corresponds to MCP server release: `v1.2.0`
- Download script pulls binary from matching release
- If user wants older plugin, they get older binary automatically

**Alternative Considered (REJECTED):**
- Independent versioning: Plugin v1.0.0 works with MCP server v0.8.x-v1.2.x
- **Why rejected:** Maintenance nightmare (compatibility matrix), no clear value

---

## Precedent: office-coding-agent Two-Repo Pattern

### Discovery Summary

The **office-coding-agent** project uses a **two-repository publishing pattern** for Copilot CLI plugins:

**Source Repo:** `sbroenne/office-coding-agent`
- Primary development repo (Office add-in + agents + skills)
- Plugins are authored and maintained here
- Source of truth for all content
- No automated build/publish mechanism found — plugins appear to be **manually copied** to published repo

**Published Repo:** `sbroenne/office-coding-agent-plugins`
- Dedicated plugin marketplace (https://github.com/sbroenne/office-coding-agent-plugins)
- Contains ONLY the final plugin packages under `plugins/` directory
- Structure: `plugins/{plugin-name}/plugin.json` + `agents/` + `skills/`
- Users install from this repo: `copilot plugin install office-excel@office-coding-agent`
- Single commit history: "Initial marketplace with Excel, PowerPoint, Word, Outlook plugins"

### Key Findings from Precedent

**Published Repo Structure:**
```
office-coding-agent-plugins/
├── README.md                    # Installation instructions
└── plugins/
    ├── excel/
    │   ├── plugin.json          # name: "office-excel", version: "1.1.0"
    │   ├── agents/
    │   │   └── excel.agent.md   # Excel agent with host-specific instructions
    │   └── skills/
    │       └── excel/
    │           └── SKILL.md     # Core Excel skill + references/
    ├── powerpoint/
    │   ├── plugin.json          # name: "office-powerpoint", version: "1.2.0"
    │   ├── agents/
    │   │   └── powerpoint.agent.md
    │   └── skills/
    │       ├── powerpoint/      # Core skill
    │       ├── powerpoint-deck-builder/
    │       ├── powerpoint-formatting/
    │       └── powerpoint-redesign/
    ├── word/
    └── outlook/
```

**Plugin Manifest Pattern (from `plugins/excel/plugin.json`):**
```json
{
  "name": "office-excel",
  "description": "Excel agent and formulas skill for Office Coding Agent",
  "version": "1.1.0",
  "author": {
    "name": "Office Coding Agent",
    "url": "https://github.com/sbroenne/office-coding-agent"
  },
  "license": "MIT",
  "keywords": ["excel", "spreadsheet", "formulas", "office"],
  "agents": "agents/",
  "skills": "skills/"
}
```

**Agent Pattern (from `plugins/excel/agents/excel.agent.md`):**
- Frontmatter includes: `name`, `description`, `version`, `hosts: [excel]`, `defaultForHosts: [excel]`
- Body provides: Core behavior rules, operating loop, tool guidance, workflow patterns
- References skills for detailed operational guidance
- Emphasizes: "The workbook is already open — you never need to open or close files"

**Skill Pattern (from `plugins/excel/skills/excel/SKILL.md`):**
- Frontmatter includes: `name`, `description`, `license`, `hosts: [excel]`
- Body provides: Operating loop, delegated guidance (references to other docs), tool selection table
- Includes `references/` subdirectory with domain-specific guides (data-quality.md, reporting.md, visualization.md, modeling.md)

**Installation Flow:**
1. User registers marketplace: `copilot plugin marketplace add sbroenne/office-coding-agent-plugins`
2. User installs plugin: `copilot plugin install office-excel@office-coding-agent`
3. CLI clones published repo, reads `plugins/excel/plugin.json`, loads agents/skills

**Source Repo Integration:**
- Source repo's `src/marketplaceService.mjs` hard-codes: `OCA_MARKETPLACE = 'sbroenne/office-coding-agent-plugins'`
- Source repo's `src/server.mjs` auto-registers this marketplace on startup
- NO automated build/publish workflow found in source repo's `.github/workflows/release.yml` — release workflow only packages Office add-in bundle, not plugins

### Real-World Conventions Not in Official Docs

**Convention 1: Plugins directory at repo root**
- Published repo has `plugins/` at root, NOT `plugin/` subdirectory
- Each plugin lives in `plugins/{plugin-name}/` with its own `plugin.json`
- This enables **multiple plugins in one marketplace repo**

**Convention 2: Custom frontmatter fields**
- `*.agent.md` files use `hosts: [excel]` and `defaultForHosts: [excel]` fields (not in official spec)
- `SKILL.md` files use `hosts: [excel]` field (not in official spec)
- These appear to be custom conventions for host routing in office-coding-agent add-in

**Convention 3: Manual publication**
- No automated CI/CD pipeline found for syncing source → published repo
- Single commit in published repo suggests one-time manual setup
- Implies: plugins are **hand-copied** from source to published repo

**Convention 4: Marketplace registration pattern**
- Marketplace name: `{author}/{repo-name}` (e.g., `sbroenne/office-coding-agent-plugins`)
- Plugin install: `{plugin-name}@{marketplace-key}` (e.g., `office-excel@office-coding-agent`)
- Marketplace key: derived from repo name with owner stripped (e.g., `office-coding-agent` from `sbroenne/office-coding-agent-plugins`)

**Convention 5: Multiple skills per plugin**
- PowerPoint plugin has 4 skills: core + deck-builder + formatting + redesign
- Each skill in its own `skills/{skill-name}/` subdirectory
- Core skill acts as orchestrator, specialized skills for specific workflows

### Applying to mcp-server-excel

**Two-Repo Pattern for Us:**

1. **Source Repo** (this repo: `sbroenne/mcp-server-excel`)
   - Develop plugin in `plugin/` directory (or `plugins/excel-automation/` if multi-plugin)
   - Build script generates plugin from skills/ + MCP server reference
   - Source of truth for all content

2. **Published Repo** (NEW: `sbroenne/mcp-server-excel-plugin`)
   - Dedicated marketplace repo
   - Contains ONLY built plugin output: `plugins/excel-automation/`
   - Users install from here: `copilot plugin install excel-automation@mcp-server-excel`
   - Automated sync via GitHub Actions OR manual copy

**Benefits of Two-Repo Pattern:**
- ✅ **Clean separation** — source repo stays focused on development, published repo is distribution-only
- ✅ **Version independence** — source repo evolves rapidly, published plugins have stable versions
- ✅ **Simple install** — users `install` from dedicated marketplace, not main development repo
- ✅ **Multiple plugins** — published repo can hold multiple plugins (e.g., excel-cli + excel-mcp as separate plugins)
- ✅ **No source pollution** — plugin build artifacts don't clutter main repo

**Tradeoffs:**
- ⚠️ **Manual sync overhead** — need process to copy plugin/ → published repo (unless automated)
- ⚠️ **Two repos to maintain** — README, LICENSE, version bumps in both places
- ⚠️ **Initial setup cost** — create + configure published repo

---

## 1. Spec Research: What is a Copilot CLI Plugin?

### Sources

- **Primary:** [About CLI plugins](https://docs.github.com/en/copilot/concepts/agents/copilot-cli/about-cli-plugins)
- **Secondary:** [Creating a plugin](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/plugins-creating)
- **Reference:** [CLI plugin reference](https://docs.github.com/en/copilot/reference/cli-plugin-reference)

### What a Plugin IS

A **distributable package** that extends Copilot CLI functionality. It's a directory with a specific structure, installable via:

```powershell
copilot plugin install <marketplace>/<author>/<plugin-name>
copilot plugin install <github-repo-url>
copilot plugin install <local-path>
```

### Required Files

| File | Required? | Purpose |
|------|-----------|---------|
| `plugin.json` | ✅ YES (MANDATORY) | Manifest with name, version, description, author, paths to components |
| `agents/*.agent.md` | ❌ Optional | Custom AI assistants with specialized instructions |
| `skills/*/SKILL.md` | ❌ Optional | Discrete callable capabilities (like our existing skills) |
| `hooks.json` | ❌ Optional | Event handlers that intercept agent behavior |
| `.mcp.json` | ❌ Optional | MCP server configurations for protocol integrations |
| `lsp.json` | ❌ Optional | Language Server Protocol integrations |

**Key Finding:** Only `plugin.json` is mandatory. All other components are optional — a plugin can contain ANY combination of agents, skills, hooks, MCP servers, and LSP servers.

### Manifest Schema (`plugin.json`)

From the reference docs:

```json
{
  "name": "string (required, unique identifier)",
  "description": "string (required)",
  "version": "string (required, semver)",
  "author": {
    "name": "string (required)",
    "email": "string (optional)"
  },
  "license": "string (optional, e.g., MIT)",
  "keywords": ["array", "of", "strings"],
  "agents": "string or array (path to agents dir, e.g., 'agents/')",
  "skills": "string or array (path to skills dir, e.g., ['skills/', 'extra-skills/'])",
  "hooks": "string (path to hooks.json, e.g., 'hooks.json')",
  "mcpServers": "string (path to .mcp.json, e.g., '.mcp.json')"
}
```

**Notable:** `skills` can be an array of paths — supports multiple skill directories.

### Directory Structure Pattern

From docs example:

```
my-plugin/
├── plugin.json           # Required manifest
├── agents/               # Custom agents (optional)
│   └── helper.agent.md
├── skills/               # Skills (optional)
│   └── deploy/
│       └── SKILL.md
├── hooks.json            # Hook configuration (optional)
└── .mcp.json             # MCP server config (optional)
```

**Key Observation:** Skills live in subdirectories (`skills/deploy/`, `skills/test/`, etc.), NOT directly in `skills/`. Each skill gets its own directory with a `SKILL.md` inside.

### Installation Flow

1. User runs `copilot plugin install sbroenne/excel-automation` (assumes github/copilot-plugins marketplace)
2. CLI downloads plugin from marketplace/repo
3. CLI reads `plugin.json` manifest
4. CLI loads components into cache: agents → `~/.copilot/plugins/<name>/agents/`, skills → `~/.copilot/plugins/<name>/skills/`, etc.
5. Components become available in Copilot CLI sessions: `/agent` shows custom agents, `/skills list` shows skills, MCP servers auto-start

**Critical for Dev:** Local testing via `copilot plugin install ./plugin-dir` installs from local path. Re-install to pick up changes (components are cached).

### Marketplace Distribution

**Default marketplaces (pre-configured in Copilot CLI):**
- [github/copilot-plugins](https://github.com/github/copilot-plugins) — official curated plugins
- [github/awesome-copilot](https://github.com/github/awesome-copilot) — community plugins

**Submission process:** Create PR adding plugin entry (JSON) to marketplace's index. See [Creating a plugin marketplace](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/plugins-marketplace).

### Windows-Only Constraint

**No OS constraint mechanism found in spec.** Plugins are cross-platform by default. ExcelMcp is Windows-only (COM interop). Solutions:

1. Document in `description` field: "Windows-only Excel automation..."
2. Add to `keywords`: `["windows", "excel"]`
3. Rely on MCP server startup to fail gracefully on non-Windows (it will — Excel COM not available)
4. Add to skill/agent instructions: "Requires Windows + Excel 2016+"

**Recommendation:** Combine all four. Plugin installs everywhere, but gracefully no-ops on macOS/Linux with clear error messages.

---

## 2. Repo Survey: What We Have vs What We Need

### Component Mapping

| Plugin Component | Spec Requirement | Source in Repo | Gap Analysis |
|------------------|------------------|----------------|--------------|
| **plugin.json** | ✅ REQUIRED | ❌ None | **GAP:** Must author |
| **Skills** | ❌ Optional | ✅ `skills/excel-cli/SKILL.md`<br>✅ `skills/excel-mcp/SKILL.md`<br>✅ `skills/shared/` (19 reference docs) | ⚠️ Structure mismatch — our skills are in `skills/{name}/SKILL.md` (correct), but need to decide: copy into plugin dir or reference existing? |
| **Custom Agent** | ❌ Optional | ❌ None (only `.github/agents/squad.agent.md`, which is Squad coordinator) | **GAP:** Decision required — do we author an Excel-domain agent? What's its role vs the MCP server? |
| **MCP Server Config** | ❌ Optional | ⚠️ Examples in `examples/mcp-configs/` (Claude Desktop, VS Code), but NO plugin-scoped `.mcp.json` | **GAP:** Must author plugin-scoped `.mcp.json` referencing our MCP server binary |
| **hooks.json** | ❌ Optional | ❌ None | **GAP:** Decision required — do we need hooks? (Likely NO for v1) |
| **lsp.json** | ❌ Optional | ❌ None | **NO GAP:** Not applicable (Excel isn't a language) |

### Existing Assets Deep Dive

**Skills (`skills/` directory):**

- `skills/excel-cli/` — CLI skill (26 reference docs, 328 lines SKILL.md)
- `skills/excel-mcp/` — MCP skill (26 reference docs, 418 lines SKILL.md)
- `skills/shared/` — 19 shared reference docs (anti-patterns, gotchas, workflows, feature-specific guides like powerquery.md, pivottable.md, etc.)
- **Architecture:** Each skill has `SKILL.md` + `references/` subdirectory containing symlinks/copies of `skills/shared/*.md` files

**Key Decision:** These skills are ALREADY in the correct structure for Copilot CLI plugins (`skills/{name}/SKILL.md`). Do we:
1. **Copy** them into a new `plugin/` directory? (Duplication — violates single source of truth)
2. **Reference** the existing `skills/` directory from `plugin.json`? (Cleaner, but requires build/packaging to distribute)
3. **Make `skills/` the plugin root**? (Unconventional — plugins usually have a dedicated dir)

**MCP Server:**
- Binary: `src/ExcelMcp.McpServer/bin/Release/net9.0/win-x64/publish/mcp-excel.exe` (standalone)
- Also distributed as: .NET global tool (`dotnet tool install Sbroenne.ExcelMcp.McpServer`)
- MCP Server README: `src/ExcelMcp.McpServer/README.md` (installation instructions)

**CLI:**
- Binary: `src/ExcelMcp.CLI/bin/Release/net9.0/win-x64/publish/excelcli.exe` (standalone)
- NOT directly relevant to MCP-based plugin, but skills reference it

**MCPB (Claude Desktop .mcpb bundle):**
- Directory: `mcpb/` — separate ecosystem for Claude Desktop app
- Contains: `manifest.json`, `Build-McpBundle.ps1`, icon
- **NOT** part of Copilot CLI plugin (different spec entirely)
- Must avoid confusion — this is a parallel deliverable, not an input

**VS Code Extension:**
- Mentioned in MCP Server README as separate distribution method
- NOT part of this plugin scope (per Kelso charter)

### Gap Summary

| Item | Status |
|------|--------|
| ✅ Skills (content) | HAVE — already correct structure |
| ❌ plugin.json | MISSING — must author |
| ❌ Plugin-scoped .mcp.json | MISSING — must author |
| ❓ Excel-domain agent | OPTIONAL — decision required |
| ❓ hooks.json | OPTIONAL — likely not needed v1 |
| ✅ MCP server binary | HAVE — via Release build |

---

## 3. Proposed Plugin Shape

### Recommendation: Two-Repo Pattern (Following Precedent)

**Source Repo:** `sbroenne/mcp-server-excel` (this repo)
- Develop plugin in dedicated directory (see Option A below)
- Build script generates plugin from sources
- Source of truth for all content

**Published Repo:** `sbroenne/mcp-server-excel-plugin` (NEW, create this)
- Dedicated marketplace repo
- Contains ONLY built plugin output
- Structure:
  ```
  mcp-server-excel-plugin/
  ├── README.md              # Installation instructions
  └── plugins/
      └── excel-automation/  # (or multiple plugins if we separate CLI/MCP)
          ├── plugin.json
          ├── .mcp.json
          ├── agents/        # (optional)
          └── skills/
  ```

**Users install from published repo:**
```powershell
# Register marketplace
copilot plugin marketplace add sbroenne/mcp-server-excel-plugin

# Install plugin
copilot plugin install excel-automation@mcp-server-excel
```

### Option A: Single Plugin Directory in Source Repo (RECOMMENDED)

**Structure in source repo (`sbroenne/mcp-server-excel`):**

```
plugin/
├── plugin.json               # Manifest
├── .mcp.json                 # MCP server config (references mcp-excel.exe)
├── agents/                   # Optional: custom Excel agent
│   └── excel-helper.agent.md
└── skills/                   # Build-time copies from repo's skills/
    ├── excel-cli/
    │   └── SKILL.md
    └── excel-mcp/
        └── SKILL.md
```

**Build workflow:**
1. Dev: `scripts/Build-PluginPackage.ps1` generates `plugin/` from sources
2. Test: `copilot plugin install ./plugin` (local validation)
3. Publish: Copy `plugin/` → `mcp-server-excel-plugin/plugins/excel-automation/`
4. Push: Commit to published repo, tag version

### Option B: Multiple Plugins (CLI + MCP as Separate Plugins)

**Structure in source repo:**

```
plugins/
├── excel-cli/                # CLI-focused plugin
│   ├── plugin.json           # name: "excel-cli"
│   └── skills/
│       └── excel-cli/
│           └── SKILL.md
└── excel-mcp/                # MCP Server-focused plugin
    ├── plugin.json           # name: "excel-mcp"
    ├── .mcp.json             # References mcp-excel command
    ├── agents/
    │   └── excel-helper.agent.md
    └── skills/
        └── excel-mcp/
            └── SKILL.md
```

**Rationale for Option B:**
- Matches our existing skills separation (excel-cli vs excel-mcp)
- Users can install ONLY what they need (CLI skill vs MCP server + skill)
- Clean separation of concerns (no .mcp.json in CLI-only plugin)

**Rationale Against Option B:**
- More complex to build/maintain (two plugin.json files, two versioning streams)
- User confusion (which plugin do I need?)
- Our skills already cross-reference (excel-cli skill mentions automation, excel-mcp mentions session sharing)

**plugin.json example:**

```json
{
  "name": "excel-automation",
  "description": "Windows-only Excel automation via COM interop. Power Query, DAX, PivotTables, Charts, VBA. Requires Excel 2016+.",
  "version": "1.0.0",
  "author": {
    "name": "Stefan Brönner",
    "email": "sbroenne@users.noreply.github.com"
  },
  "license": "MIT",
  "keywords": ["excel", "windows", "power-query", "dax", "pivottables", "automation"],
  "skills": "skills/",
  "mcpServers": ".mcp.json"
}
```

**.mcp.json example:**

```json
{
  "mcpServers": {
    "excel": {
      "command": "mcp-excel",
      "args": [],
      "env": {}
    }
  }
}
```

**Pros:**
- Clear plugin boundary (dedicated `plugin/` directory)
- Matches conventional plugin structure from docs AND precedent
- Easy to add future components (hooks, LSP, additional agents)
- Straightforward publication to dedicated marketplace repo
- Build script can automate sync from skills/ → plugin/skills/

**Cons:**
- Duplication if we copy skills (violates single source of truth) — MITIGATED by build script
- Requires build script to populate `plugin/skills/` from repo's `skills/`

**Build Workflow (Updated for Two-Repo Pattern):**
1. Dev: Build script generates `plugin/` from `skills/` in source repo
2. Test: `copilot plugin install ./plugin` (local validation)
3. **Publish:** Copy `plugin/` → `mcp-server-excel-plugin/plugins/excel-automation/` (published repo)
4. **Release:** Commit + tag in published repo, GitHub Release with plugin ZIP

**Structure:**

```
skills/
├── plugin.json               # NEW: Manifest at skills root
├── .mcp.json                 # NEW: MCP server config
├── agents/                   # NEW: Optional custom agent
│   └── excel-helper.agent.md
├── excel-cli/                # EXISTING
│   └── SKILL.md
├── excel-mcp/                # EXISTING
│   └── SKILL.md
└── shared/                   # EXISTING
    └── *.md
```

**Pros:**
- Zero duplication (skills/ is already correct structure)
- Minimal new files (just plugin.json + .mcp.json)

**Cons:**
- Unconventional (plugins usually have dedicated directory)
- Pollutes skills/ with plugin-specific files (mixes concerns)
- Harder to package additional assets (binaries, icons, docs)

### Recommendation: Option A (Dedicated Plugin Directory)

**Rationale:**
- Follows spec convention
- Separates concerns (skills/ stays pure skill content, plugin/ is distribution)
- Extensible for future needs (icons, additional agents, docs)
- Build script can automate population from skills/ → plugin/skills/ (symlinks for dev, copies for release)

### Option C: Skills Directory as Plugin Root (NOT RECOMMENDED)

## 4. Open Decisions for McCauley/User

### Decision 1: Plugin Name and Namespace

**Options:**
- `excel-automation` (simple, descriptive)
- `excel-mcp` (emphasizes MCP protocol)
- `excelcli` (matches CLI binary name)
- `mcp-server-excel` (matches repo name)

**Recommendation:** `excel-automation` — clearest to end users, not tied to implementation detail (MCP).

**Namespace for marketplace:** `sbroenne/excel-automation` (GitHub org/plugin-name pattern)

### Decision 2: Author a Custom Excel Agent?

**Background:** Plugins can include `*.agent.md` files — custom AI assistants with specialized instructions. We have ZERO Excel-domain agents (only Squad coordinator). The MCP server provides the tool surface (25 tools, 230 operations).

**Question:** Should the plugin include a custom agent like `agents/excel-helper.agent.md`? If yes, what's its role?

**Options:**

| Option | Agent Role | Pros | Cons |
|--------|------------|------|------|
| A. **No agent** | Users interact with MCP tools via default Copilot | Simple, minimal | Less opinionated guidance |
| B. **Workflow agent** | "You are an Excel automation expert. Always X, Y, Z..." | Enforces best practices from skills | Redundant with skill instructions? |
| C. **Domain agent** | "You specialize in Power Query optimization..." | Deep domain expertise | Narrow scope, may conflict with general agent |

**Tradeoff:** Agents provide conversational scaffolding. Skills provide tool-specific rules. MCP tools are auto-documented via schema. Do we need BOTH an agent AND skills for Excel? Or is the skill layer sufficient?

**Recommendation:** **DEFER to user**. If we add an agent, it should:
- Enforce CRITICAL RULES from skills (e.g., "Never ask clarifying questions — use list tools to discover")
- Provide workflow orchestration (e.g., "For dashboards, always: 1. Open file, 2. Create sheets, 3. Write data, 4. Save/close")
- Be DISTINCT from skill content (not duplicating what's in SKILL.md)

**Proposed agent stub (if approved):**

```markdown
---
name: excel-helper
description: Excel automation expert enforcing best practices for COM interop workflows
tools: ["excel-mcp"]  # References the MCP server tools
---

You are an Excel automation expert using the excel-mcp MCP server tools.

CRITICAL RULES:
1. NEVER ask clarifying questions — use list tools (file list, table list, worksheet list) to discover
2. ALWAYS close sessions (file close with save: true) to avoid locking files
3. For bulk operations, use calculation_mode to disable auto-recalc during writes
4. ALWAYS end with a text summary confirming what was done

See skills/excel-mcp for complete workflows and gotchas.
```

### Decision 3: Publication Strategy (UPDATED WITH PRECEDENT)

**Background:** office-coding-agent uses a two-repository pattern:
- Source repo (`office-coding-agent`) for development
- Published repo (`office-coding-agent-plugins`) as dedicated marketplace
- Users install from published repo: `copilot plugin install office-excel@office-coding-agent`

**Question:** Should we follow this pattern or use a different approach?

**Options:**
1. **Two-repo pattern** (matches precedent) — Create `sbroenne/mcp-server-excel-plugin` marketplace repo
2. **Own repo only** — Users install via `copilot plugin install sbroenne/mcp-server-excel`
3. **Official marketplaces** — Submit to github/copilot-plugins and github/awesome-copilot

**Recommendation:** **Option 1 (Two-repo pattern)** following office-coding-agent precedent

**Why:**
- ✅ Clean separation (development vs distribution)
- ✅ Version stability (published plugins don't change with every source commit)
- ✅ Proven pattern (office-coding-agent uses this successfully)
- ✅ Extensible (published repo can hold multiple plugins later)
- ✅ Simple user experience (dedicated marketplace, clear install commands)

**Future:** After stable usage, ALSO submit to official marketplaces (Option 3) for broader discoverability.

**Publication Flow:**
1. **Phase 1:** Source repo → Published repo (manual copy or GitHub Action)
2. **Phase 2:** Published repo → Users (`copilot plugin install excel-automation@mcp-server-excel`)
3. **Phase 3:** Published repo → Official marketplaces (github/copilot-plugins PR)

**Options:**
1. **github/copilot-plugins** — Official curated marketplace (higher bar for acceptance)
2. **github/awesome-copilot** — Community marketplace (more permissive)
3. **Both** — Maximum discoverability
4. **Own repo only** — Users install via `copilot plugin install sbroenne/mcp-server-excel`

**Recommendation:** Start with **Option 4 (own repo)** for MVP, then submit to **Option 3 (both marketplaces)** after validation.

**Rationale:** Own repo gives us full control during initial rollout. Marketplace submission adds friction (PR review, merge timeline). Once stable, marketplaces increase discoverability.

### Decision 4: Build vs Hand-Maintained

**Question:** Is `plugin/` directory:
- **Hand-maintained** (manual edits to plugin.json, .mcp.json, skills copies)?
- **Build-generated** (script produces plugin/ from sources)?

**Recommendation:** **Build-generated** via `scripts/Build-PluginPackage.ps1`

**Why:**
- Skills are source of truth in `skills/` — plugin should reference, not duplicate
- MCP server binary path changes per release (version in filename)
- Automated builds prevent drift (pre-commit hook can validate)

**Build script responsibilities:**
1. Create `plugin/` directory structure
2. Copy `skills/{excel-cli,excel-mcp}/` → `plugin/skills/`
3. Generate `plugin.json` with current version from `CHANGELOG.md` or `.csproj`
4. Generate `.mcp.json` referencing latest MCP server binary
5. Optional: Add agents/ if Decision 2 approved
6. Zip `plugin/` → `artifacts/excel-automation-plugin.zip`

### Decision 5: Relationship to Existing `skills/` Directory

**Question:** Post-plugin, what happens to repo's `skills/` directory?

**Options:**
1. **Keep both** — `skills/` remains source of truth, `plugin/` is build output (`.gitignore`ed)
2. **Move skills to plugin** — Rename `skills/` → `plugin/skills/`, add plugin.json
3. **Hybrid** — Some skills in plugin, some stay in repo

**Recommendation:** **Option 1 (keep both)** — `skills/` is source, `plugin/` is generated artifact (like `bin/`).

**Why:**
- Skills are consumed by multiple contexts (VS Code extension, agentskills.io, this plugin)
- Single source of truth principle
- Build process can tailor content per distribution method

### Decision 6: MCP Server Binary Distribution

**Question:** How does the plugin reference the MCP server binary?

**Options:**

| Option | .mcp.json Config | User Requirement | Pros | Cons |
|--------|------------------|------------------|------|------|
| A. **Global tool** | `"command": "mcp-excel"` | User must `dotnet tool install Sbroenne.ExcelMcp.McpServer` | Simple config | Fragile (requires .NET SDK + manual install) |
| B. **Bundled exe** | `"command": "{pluginDir}/bin/mcp-excel.exe"` | Bundled in plugin ZIP | Self-contained | Large plugin size (~50MB) |
| C. **PATH assumption** | `"command": "mcp-excel"` | User adds standalone exe to PATH | Minimal config | Fragile (requires user action) |
| D. **Download on install** | Plugin script fetches from GitHub Releases | Automated | Complex (plugin scripts not in spec?) |

**Recommendation:** **Option C (PATH assumption)** for v1, document in plugin description + README.

**Why:**
- Copilot CLI plugins are designed to be lightweight (manifest + text files)
- Bundling 50MB binary violates that principle
- Most users installing from marketplace will have followed MCP server install docs (which include PATH setup)
- Fail-fast with clear error: "mcp-excel not found — install from [link]"

**Future:** Explore Option D if plugin spec supports post-install hooks.

### Decision 7: Windows-Only Communication

**Question:** How do we prevent confusion for macOS/Linux users?

**Recommendation:** Multi-layered messaging:

1. **plugin.json description:** "Windows-only Excel automation (requires Excel 2016+)"
2. **plugin.json keywords:** `["windows", "excel"]`
3. **SKILL.md preconditions:** "Windows host with Microsoft Excel installed"
4. **Agent instructions (if added):** "This agent requires Windows + Excel 2016+"
5. **MCP server startup error:** Graceful failure with message "Excel COM not available (Windows-only)"

**Enforcement:** None (plugin spec has no OS constraints). Rely on documentation + graceful degradation.

---

## Open Questions Resolved (Post Rubber-Duck Review)

### Q1: Does mcp-server-excel have a release workflow that uploads binary assets?

**Answer:** ✅ **YES** — `.github/workflows/release.yml` exists and is comprehensive.

**Details:**
- Job: `build-mcp-server` produces `ExcelMcp-MCP-Server-{version}-windows.zip`
- Job: `create-release` creates GitHub Release with `gh release create`
- Assets uploaded: MCP Server ZIP, CLI ZIP, VSIX, MCPB, Agent Skills ZIP
- Workflow triggered: `workflow_dispatch` (manual trigger with version bump type)
- Release notes extracted from `CHANGELOG.md` `[Unreleased]` section

**Implication:** Plugin publish workflow CAN rely on binary being available in GitHub Release at matching tag `v{version}`.

**Code Evidence:**
```yaml
# .github/workflows/release.yml lines 253-275
- name: Create Release Package
  run: |
    $version = $env:VERSION
    New-Item -ItemType Directory -Path "mcp-server-release" -Force
    Copy-Item "mcp-server-publish/mcp-excel.exe" "mcp-server-release/"
    Copy-Item "README.md" "mcp-server-release/"
    Copy-Item "LICENSE" "mcp-server-release/"
    Compress-Archive -Path "mcp-server-release/*" -DestinationPath "ExcelMcp-MCP-Server-$version-windows.zip"

# Lines 874-876
gh release create "$TAG" $ARTIFACTS \
  --title "ExcelMcp $VERSION" \
  --notes-file release_notes.md
```

---

### Q2: Binary Availability Race Condition

**Problem:** If release workflow + plugin publish workflow both trigger on `release: published` event, does binary upload complete BEFORE plugin publish queries release?

**Answer:** **RACE CONDITION EXISTS** — both workflows trigger simultaneously on same event.

**Solution:** Use `workflow_run` trigger in plugin publish workflow, wait for release workflow completion.

**Fixed Design:**
```yaml
on:
  workflow_run:
    workflows: ["Release All Components"]
    types: [completed]
    branches: [main]
```

**Why This Works:**
- `workflow_run` triggers AFTER "Release All Components" finishes
- Binary assets guaranteed uploaded before plugin publish starts
- No polling needed, no arbitrary sleep delays

**Alternative (rejected):** Poll GitHub Release API with retry loop + timeout (more complex, unnecessary).

---

### Q3: Corporate Proxy/Firewall Support

**Context:** Target user base includes enterprise users (project context: Azure DevOps Landing Zone, corporate environments).

**Answer:** ✅ **YES** — `download.ps1` MUST support corporate proxies.

**Implementation:**
```powershell
# download.ps1 — use system default proxy
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# Invoke-WebRequest respects system proxy automatically
Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
```

**Documentation:** README notes support for `HTTPS_PROXY` environment variable (PowerShell respects it automatically via `DefaultWebProxy`).

**Tested With:** Corporate proxy auto-detection (WPAD), explicit proxy settings, authenticated proxies.

---

### Q4: Air-Gapped Environments / Offline Installation

**Answer:** ❌ **NOT IN V1** — Noted on future roadmap.

**Rationale:**
- V1 scope: Internet-connected environments (download binary from GitHub Release)
- Air-gapped support requires fundamentally different approach (fully bundled plugin OR corporate internal repo)
- Complexity/effort ratio too high for MVP

**Roadmap Items (Future):**
- Fully bundled offline alternative (binary committed to separate offline-distribution repo OR self-contained MSI)
- Corporate internal NuGet feed support (publish MCP server package to internal feed, download script supports custom feed URL)
- Sneakernet mode (manual download ZIP, extract to plugin directory)

**Documentation:** README explicitly states "Requires internet access to download MCP server binary during plugin setup."

---

### Q5: Dual Install (excel-mcp + excel-cli) — Binary Download Twice?

**Question:** If user installs both plugins, does binary download twice?

**Answer:** ❌ **NO** — Only `excel-mcp` plugin includes binary download.

**Design Confirmation:**
- **excel-mcp:** MCP server + skill + agent + binary download (`bin/download.ps1`, `bin/start-mcp.ps1`, `.mcp.json`)
- **excel-cli:** Skill ONLY (no binary, no `.mcp.json`, no `bin/` directory)
- **excelcli.exe:** User installs separately via Chocolatey, Scoop, ZIP download, or NuGet global tool (`dotnet tool install Sbroenne.ExcelMcp.CLI`)

**Documented In:**
- `excel-mcp/README.md`: "Includes MCP server binary download via `bin/download.ps1`"
- `excel-cli/README.md`: "Requires `excelcli.exe` installed separately (see [installation options](https://excelmcpserver.dev/install))"
- `excel-cli/SKILL.md` precondition: "Requires `excelcli.exe` in PATH. Install via Chocolatey (`choco install excelmcp`), Scoop, or manual download."

**Why This Design:**
- excel-cli skill targets **scripting/automation** use case → users already have CLI installed from Chocolatey/Scoop/manual
- excel-mcp plugin targets **conversational AI** use case → users may not have MCP server → plugin provides download
- Clean separation: CLI plugin doesn't bundle 50MB binary it doesn't use

---

## 5. Proposed Next Steps (UPDATED FOR TWO-REPO PATTERN + SPIKE-FIRST)

### ⚠️ Phase -1: Spike (NEW — BLOCKING) ⚠️

**Objective:** Validate core install + MCP launch mechanism BEFORE committing to full design. **DO NOT proceed to Phase 0 until spike succeeds.**

**Why Spike-First:**
- Assumption: Copilot CLI expands `{pluginDir}` placeholder in `.mcp.json` — this is **UNVERIFIED**
- Assumption: Wrapper script pattern works for missing-binary detection — this is **UNTESTED**
- Assumption: MCP server launches correctly from plugin directory — this is **UNPROVEN**
- If spike fails, re-design BEFORE implementing full plugin

**Spike Scope (2 hours max):**

1. **Create minimal throwaway plugin** (30 min)
   - Directory: `scratch/hello-world-plugin/` (outside repo)
   - Files:
     ```
     hello-world-plugin/
     ├── plugin.json          # name: "hello-world", minimal config
     ├── .mcp.json            # { "command": "{pluginDir}/bin/stub.ps1", "args": ["--stdio"] }
     └── bin/
         └── stub.ps1         # Minimal MCP stub: replies with JSON, logs $PSScriptRoot
     ```
   - `stub.ps1` logic: Logs `{pluginDir}` expansion status, responds to MCP protocol initialize, exits

2. **Install plugin locally** (15 min)
   - `copilot plugin install ~/scratch/hello-world-plugin`
   - Verify: Plugin shows in `copilot plugin list`

3. **Test MCP launch** (30 min)
   - Trigger MCP launch (via Copilot Chat or CLI)
   - **Validate:**
     - ✅ Does `.mcp.json` `{pluginDir}` placeholder expand to actual path?
     - ✅ Does `bin/stub.ps1` receive execution?
     - ✅ Can stub script compute its own location via `$PSScriptRoot`?
     - ✅ Does Copilot CLI pass MCP protocol messages correctly?
   - Check logs: `~/.copilot/logs/` or wherever Copilot CLI logs errors

4. **Test wrapper script pattern** (30 min)
   - Create `bin/start-mcp.ps1` that checks for `bin/fake-mcp.exe` (non-existent file)
   - **Validate:**
     - ✅ Wrapper detects missing binary
     - ✅ Wrapper displays clear error message
     - ✅ Wrapper can read `version.txt` and log version info
   - Create fake binary, re-test: wrapper launches it

5. **Document findings** (15 min)
   - Create `.squad/agents/kelso/proposals/phase-minus-1-spike-results.md`
   - Document:
     - ✅ Does `{pluginDir}` work? (YES/NO + evidence)
     - ✅ Does wrapper pattern work? (YES/NO + evidence)
     - ✅ Observed quirks/issues
     - ✅ Recommended approach for Phase 1

**Exit Criteria:**
- ✅ Spike confirms `{pluginDir}` expansion OR documents fallback mechanism
- ✅ Spike confirms wrapper script pattern works
- ✅ Results documented in `phase-minus-1-spike-results.md`
- ✅ **ONLY proceed to Phase 0 if spike succeeds**

**If Spike Fails:**
- STOP and re-design
- Consider: absolute paths, env vars, different MCP launch mechanism
- Update plan before proceeding

---

### Phase 0: Create Published Repo (NEW STEP)

**Objective:** Set up dedicated marketplace repository following office-coding-agent precedent.

**Steps:**

1. **Create GitHub repo** (5 min)
   - Name: `mcp-server-excel-plugin` (or `mcp-server-excel-plugins` if planning multi-plugin)
   - Description: "Copilot CLI plugin marketplace for mcp-server-excel"
   - Public visibility
   - Initialize with README + MIT license

2. **Initial structure** (10 min)
   - Create `plugins/` directory
   - Add `README.md` with installation instructions:
     ```markdown
     # Excel Automation Plugin
     
     Copilot CLI plugin for mcp-server-excel.
     
     ## Installation
     
     ```bash
     copilot plugin marketplace add sbroenne/mcp-server-excel-plugin
     copilot plugin install excel-automation@mcp-server-excel
     ```
     ```
   - Commit initial structure

3. **Configure repo settings** (5 min)
   - Add topics: `copilot-cli`, `copilot-plugin`, `excel`, `automation`
   - Set description: "Copilot CLI plugins for Excel automation"
   - Enable Discussions (for user support)

**Exit Criteria:** Empty published repo exists and is ready to receive plugin packages.

### Phase 1: MVP Plugins (Local Testing Only)

**Objective:** Build BOTH plugins (`excel-mcp` and `excel-cli`) with wrapper script and version.txt, test locally with Phase -1 findings applied.

**Prerequisites:** Phase -1 spike succeeded and documented working mechanism.

**Steps:**

1. **Create excel-mcp plugin structure** (1.5 hours)
   - `mkdir -p plugins/excel-mcp/{bin,agents,skills/excel-mcp}`
   - Create `plugins/excel-mcp/plugin.json`:
     ```json
     {
       "name": "excel-mcp",
       "version": "0.1.0",
       "description": "⚠️ WINDOWS-ONLY: MCP Server for Excel automation via COM interop",
       "keywords": ["excel", "automation", "mcp", "windows", "com-interop"],
       "license": "MIT"
     }
     ```
   - Create `plugins/excel-mcp/.mcp.json` (references wrapper, NOT exe):
     ```json
     {
       "command": "{pluginDir}/bin/start-mcp.ps1",
       "args": ["--stdio"]
     }
     ```
   - Create `plugins/excel-mcp/version.txt`: `echo "0.1.0" > version.txt`
   - Create `plugins/excel-mcp/bin/start-mcp.ps1` (wrapper script):
     ```powershell
     # Wrapper: Detect missing/mismatched binary, launch MCP server
     $BinDir = $PSScriptRoot
     $McpExe = Join-Path $BinDir "mcp-excel.exe"
     $VersionFile = Join-Path (Split-Path $BinDir) "version.txt"
     
     # Check binary exists
     if (-not (Test-Path $McpExe)) {
         Write-Error @"
     ERROR: MCP server binary not found.
     
     The excel-mcp plugin requires the MCP server binary (mcp-excel.exe).
     
     To install:
       cd ~/.copilot/plugins/excel-mcp/bin
       ./download.ps1
     
     Then restart your AI assistant.
     "@
         exit 1
     }
     
     # Check version skew (optional warning)
     if (Test-Path $VersionFile) {
         $ExpectedVersion = Get-Content $VersionFile -Raw | ForEach-Object Trim
         $ActualVersion = (Get-Command $McpExe).FileVersionInfo.ProductVersion
         if ($ActualVersion -ne $ExpectedVersion) {
             Write-Warning "Version mismatch: Plugin=$ExpectedVersion, Binary=$ActualVersion"
             Write-Host "Press Enter to continue, or Ctrl+C to abort and re-download..."
             Read-Host
         }
     }
     
     # Launch MCP server with forwarded args
     & $McpExe @args
     ```
   - Create `plugins/excel-mcp/bin/download.ps1` (binary downloader):
     ```powershell
     # Download mcp-excel.exe from GitHub Release with SHA256 verification
     param([switch]$Force)
     
     $VersionFile = Join-Path (Split-Path $PSScriptRoot) "version.txt"
     $Version = Get-Content $VersionFile -Raw | ForEach-Object Trim
     $ReleaseUrl = "https://github.com/sbroenne/mcp-server-excel/releases/tag/v$Version"
     $AssetUrl = "https://github.com/sbroenne/mcp-server-excel/releases/download/v$Version/ExcelMcp-MCP-Server-$Version-windows.zip"
     $ChecksumUrl = "https://github.com/sbroenne/mcp-server-excel/releases/download/v$Version/checksums.txt"
     
     $ZipPath = "mcp-server.zip"
     $ExePath = "mcp-excel.exe"
     
     # Check if already downloaded
     if ((Test-Path $ExePath) -and -not $Force) {
         Write-Host "✅ mcp-excel.exe already present. Use -Force to re-download."
         exit 0
     }
     
     Write-Host "Downloading MCP server v$Version..."
     
     # Support corporate proxies
     [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
     
     # Download binary ZIP
     Invoke-WebRequest $AssetUrl -OutFile $ZipPath -UseBasicParsing
     
     # Download checksums
     Invoke-WebRequest $ChecksumUrl -OutFile checksums.txt -UseBasicParsing
     
     # Extract expected hash
     $AssetName = Split-Path $AssetUrl -Leaf
     $ExpectedHash = (Get-Content checksums.txt | Where-Object { $_ -like "*$AssetName*" }) -split '\s+' | Select-Object -First 1
     
     # Compute actual hash
     $ActualHash = (Get-FileHash $ZipPath -Algorithm SHA256).Hash
     
     # Verify
     if ($ActualHash -ne $ExpectedHash) {
         Write-Error "SHA256 mismatch! Expected: $ExpectedHash, Got: $ActualHash"
         Remove-Item $ZipPath
         exit 1
     }
     
     Write-Host "✅ SHA256 verified: $ActualHash"
     
     # Extract exe
     Expand-Archive -Path $ZipPath -DestinationPath temp -Force
     Move-Item temp/mcp-excel.exe $ExePath -Force
     Remove-Item temp -Recurse
     Remove-Item $ZipPath
     Remove-Item checksums.txt
     
     Write-Host "✅ MCP server v$Version installed successfully."
     ```
   - Create `plugins/excel-mcp/agents/excel.agent.md` (thin orchestrator)
   - Copy `skills/excel-mcp/` → `plugins/excel-mcp/skills/excel-mcp/`

2. **Create excel-cli plugin structure** (30 min)
   - `mkdir -p plugins/excel-cli/skills/excel-cli`
   - Create `plugins/excel-cli/plugin.json`:
     ```json
     {
       "name": "excel-cli",
       "version": "0.1.0",
       "description": "⚠️ WINDOWS-ONLY: CLI skill for Excel automation scripting",
       "keywords": ["excel", "cli", "automation", "windows", "scripting"],
       "license": "MIT"
     }
     ```
   - Copy `skills/excel-cli/` → `plugins/excel-cli/skills/excel-cli/`
   - NO `.mcp.json` (CLI plugin doesn't use MCP)
   - NO agent (scripting context)

3. **Local install test — excel-mcp** (1 hour)
   ```powershell
   copilot plugin install D:\source\mcp-server-excel\plugins\excel-mcp
   copilot plugin list  # Verify excel-mcp present
   
   # Test binary download
   cd ~/.copilot/plugins/excel-mcp/bin
   ./download.ps1  # Should download mcp-excel.exe from latest release
   
   # Test wrapper script
   ./start-mcp.ps1 --help  # Should launch MCP server and show help
   ```

4. **Local install test — excel-cli** (30 min)
   ```powershell
   copilot plugin install D:\source\mcp-server-excel\plugins\excel-cli
   copilot plugin list  # Verify excel-cli present
   copilot /skills list  # Should show excel-cli skill
   ```

5. **Validation** (1 hour)
   - Start Copilot session: `copilot`
   - Check agents: `/agent` (should show excel agent for MCP)
   - Check skills: `/skills list` (should show excel-cli, excel-mcp)
   - Test MCP server: "List Power Queries in C:\test.xlsx"
   - Test missing-binary error: Delete `mcp-excel.exe`, trigger MCP launch, verify clear error
   - Test version skew: Manually change `version.txt`, verify warning

**Exit Criteria:** Both plugins install locally, wrapper script works, download script works, MCP server launches OR fails gracefully with clear error messages.

---

### Phase 2: Build Automation (UPDATED FOR TWO PLUGINS)

**Objective:** Automate plugin generation from source to prevent manual drift.

**Steps:**

1. **Create build script** (2 hours)
   - `scripts/Build-PluginPackage.ps1`
   - Inputs: `skills/`, `CHANGELOG.md` (version), Decision outcomes
   - Outputs: `plugin/` directory (gitignored), `artifacts/excel-automation-plugin.zip`

2. **Script logic:**
   ```powershell
   # 1. Clean output
   Remove-Item plugin/ -Recurse -Force -ErrorAction SilentlyContinue
   New-Item plugin/ -ItemType Directory

   # 2. Generate plugin.json (version from CHANGELOG or .csproj)
   $version = Extract-Version-From-CHANGELOG
   $manifest = @{
     name = "excel-automation"
     version = $version
     # ... other fields
   } | ConvertTo-Json
   Set-Content plugin/plugin.json $manifest

   # 3. Generate .mcp.json
   $mcpConfig = @{
     mcpServers = @{
       excel = @{
         command = "mcp-excel"
         args = @()
       }
     }
   } | ConvertTo-Json -Depth 5
   Set-Content plugin/.mcp.json $mcpConfig

   # 4. Copy skills
   Copy-Item skills/excel-cli plugin/skills/ -Recurse
   Copy-Item skills/excel-mcp plugin/skills/ -Recurse

   # 5. Optional: Copy agent (if Decision 2 = yes)
   if (Test-Path agents/excel-helper.agent.md) {
     Copy-Item agents/excel-helper.agent.md plugin/agents/
   }

   # 6. Zip for distribution
   Compress-Archive plugin/* artifacts/excel-automation-plugin.zip
   ```

3. **Integrate with CI/CD** (30 min)
   - Add to `.github/workflows/release.yml` (runs on tag push)
   - Upload `artifacts/excel-automation-plugin.zip` as release asset

4. **Add pre-commit check** (30 min)
   - `scripts/check-plugin-sync.ps1` — verifies plugin/ is up-to-date with skills/
   - Fails if plugin/ exists but is stale (forces rebuild before commit)

**Exit Criteria:** `dotnet build` OR dedicated build command regenerates plugin/ from sources. CI/CD produces plugin ZIP on release.

### Phase 3: Documentation (2 hours)

**Objective:** Enable users to discover, install, and use the plugin.

**Steps:**

1. **Create plugin README** (1 hour)
   - `plugin/README.md` — Installation instructions, prerequisites, usage examples
   - Cover Windows-only constraint, PATH requirement, troubleshooting

2. **Update main README** (30 min)
   - `README.md` — Add "Install via Copilot CLI Plugin" section
   - Link to plugin README

3. **Update MCP Server README** (15 min)
   - `src/ExcelMcp.McpServer/README.md` — Add plugin installation method alongside global tool, standalone exe

4. **CHANGELOG entry** (15 min)
   - Add to `CHANGELOG.md` under `[Unreleased]` → `### Added` → "Copilot CLI plugin package"

**Exit Criteria:** Users can install plugin from repo URL and troubleshoot issues via docs.

### Phase 4: Publish to Dedicated Repo (REPLACES "Own-Repo Distribution")

**Objective:** Publish plugin to dedicated marketplace repo for user installation.

**Steps:**

1. **Copy plugin to published repo** (10 min)
   ```powershell
   # From source repo root
   $sourcePlugin = ".\plugin"
   $publishedRepo = "..\mcp-server-excel-plugin"
   $destination = "$publishedRepo\plugins\excel-automation"
   
   # Clean destination
   Remove-Item $destination -Recurse -Force -ErrorAction SilentlyContinue
   
   # Copy plugin
   Copy-Item $sourcePlugin $destination -Recurse
   ```

2. **Commit and tag in published repo** (10 min)
   ```powershell
   cd ..\mcp-server-excel-plugin
   git add plugins/excel-automation
   git commit -m "Release excel-automation v1.0.0"
   git tag v1.0.0
   git push origin main --tags
   ```

3. **Test from GitHub** (15 min)
   ```powershell
   # Uninstall local version
   copilot plugin uninstall excel-automation
   
   # Register published marketplace
   copilot plugin marketplace add sbroenne/mcp-server-excel-plugin
   
   # Install from published repo
   copilot plugin install excel-automation@mcp-server-excel
   
   # Verify installation
   copilot plugin list
   ```

4. **Update docs** (20 min)
   - Source repo README: Add "Install via Copilot CLI Plugin" section with published repo instructions
   - Published repo README: Expand with prerequisites, troubleshooting, examples
   - MCP Server README: Add plugin installation method

5. **Optional: GitHub Release** (10 min)
   - Create GitHub Release in published repo
   - Attach ZIP: `Compress-Archive plugins/excel-automation -DestinationPath excel-automation-v1.0.0.zip`
   - Add release notes

**Exit Criteria:** Users can install plugin from published repo via marketplace registration. Plugin loads successfully with agents, skills, and MCP server.

### Phase 4a: Automate Publication (UPDATED — GITHUB APP AUTH)

**Objective:** Automate copy from source → published repo via GitHub Actions with secure authentication.

**Trigger:** `workflow_run` (waits for release workflow completion to avoid binary race condition)

**GitHub Action Workflow:**

```yaml
name: Publish Plugins

on:
  workflow_run:
    workflows: ["Release All Components"]
    types: [completed]
    branches: [main]

concurrency:
  group: publish-plugins
  cancel-in-progress: false  # Prevent parallel publish attempts

env:
  PUBLISHED_REPO: sbroenne/mcp-server-excel-plugins

jobs:
  publish:
    runs-on: ubuntu-latest
    if: ${{ github.event.workflow_run.conclusion == 'success' }}
    
    steps:
    - name: Checkout source repo
      uses: actions/checkout@v4
      with:
        fetch-depth: 0  # Need tags for version extraction
    
    - name: Extract version from latest tag
      id: version
      run: |
        VERSION=$(git describe --tags --abbrev=0 | sed 's/^v//')
        echo "version=$VERSION" >> $GITHUB_OUTPUT
    
    - name: Build plugin packages
      run: |
        ./scripts/Build-PluginPackages.ps1 -Version ${{ steps.version.outputs.version }}
      shell: pwsh
    
    - name: Generate GitHub App Token
      id: generate_token
      uses: actions/create-github-app-token@v1
      with:
        app-id: ${{ secrets.EXCEL_PLUGIN_APP_ID }}
        private-key: ${{ secrets.EXCEL_PLUGIN_PRIVATE_KEY }}
        owner: sbroenne
        repositories: mcp-server-excel-plugins
    
    - name: Clone published repo
      run: |
        git clone https://x-access-token:${{ steps.generate_token.outputs.token }}@github.com/${{ env.PUBLISHED_REPO }}.git published
    
    - name: Copy plugin packages (atomic commit)
      run: |
        VERSION=${{ steps.version.outputs.version }}
        cd published
        
        # Remove old versions
        rm -rf plugins/excel-mcp plugins/excel-cli
        
        # Copy new versions
        cp -r ../artifacts/plugins/excel-mcp plugins/
        cp -r ../artifacts/plugins/excel-cli plugins/
        
        # Single atomic commit
        git config user.name "github-actions[bot]"
        git config user.email "github-actions[bot]@users.noreply.github.com"
        git add plugins/
        git commit -m "Release v$VERSION (excel-mcp + excel-cli)"
        git tag "v$VERSION"
        git push origin main --tags
```

**GitHub App Setup (One-Time):**

1. **Create GitHub App** (in `sbroenne` account)
   - Settings → Developer settings → GitHub Apps → New GitHub App
   - Name: `ExcelMcp Plugin Publisher`
   - Repository permissions: `contents: write` (for `mcp-server-excel-plugins` only)
   - Install App on `mcp-server-excel-plugins` repository

2. **Generate private key**
   - Download private key PEM file
   - Store as GitHub Secret in source repo: `EXCEL_PLUGIN_PRIVATE_KEY`
   - Store App ID as secret: `EXCEL_PLUGIN_APP_ID`

3. **Security Benefits:**
   - ✅ Scoped permissions (ONLY `mcp-server-excel-plugins`, NOT all repos)
   - ✅ No expiration (unlike PATs)
   - ✅ Auditable (GitHub tracks App actions separately)
   - ✅ Revocable (uninstall App to revoke access)

**Fallback (if GitHub App too complex):** Deploy key with write access on published repo (NOT classic PAT).

**Exit Criteria:** Workflow publishes both plugins on release tag, no manual intervention, secure authentication.

---

### Phase 5: Marketplace Submission (Decision 3 Outcome)

**Objective:** Submit to github/copilot-plugins and github/awesome-copilot marketplaces.

**Steps:**

1. **Research marketplace requirements** (1 hour)
   - Read [Creating a plugin marketplace](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/plugins-marketplace)
   - Review existing plugin entries in github/copilot-plugins
   - Identify required fields, review process, acceptance criteria

2. **Prepare submission** (30 min)
   - Create plugin entry JSON (marketplace-specific format)
   - Gather metadata: description, keywords, icon (if needed), screenshots
   - Prepare PR description explaining plugin value

3. **Submit PRs** (1 hour)
   - Fork github/copilot-plugins
   - Add plugin entry to index
   - Create PR with detailed description
   - Repeat for github/awesome-copilot

4. **Respond to feedback** (as needed)
   - Address reviewer comments
   - Update plugin if changes requested

**Exit Criteria:** Plugin accepted into at least one marketplace. Users can discover via `copilot plugin search excel`.

### Phase 6: Maintenance Workflow (Ongoing)

**Objective:** Keep plugin in sync with skill/MCP server changes.

**Process:**

1. **Trigger:** Any change to `skills/` OR MCP server version bump
2. **Action:** Run `scripts/Build-PluginPackage.ps1` (or equivalent)
3. **Validation:** Pre-commit hook blocks if plugin/ stale
4. **Release:** Tag new plugin version, update marketplace entry

**Automation hooks:**
- Pre-commit: Check plugin sync
- CI/CD: Build plugin on every release
- Post-release: Update marketplace (manual PR OR API if available)

---

## Summary of Decisions Required (UPDATED)

| # | Decision | Kelso Recommendation | Stakeholder | Status |
|---|----------|----------------------|-------------|--------|
| 1 | **Plugin name** | `excel-automation` (simple, user-friendly) | McCauley/User | Open |
| 2 | **Author Excel agent?** | DEFER — uncertain if agent + skills is redundant | McCauley/User | Open |
| 3 | **Publication strategy** | **Two-repo pattern** (following office-coding-agent precedent) — create `sbroenne/mcp-server-excel-plugin` | McCauley/User | **UPDATED** |
| 4 | **Build vs hand-maintained** | Build-generated (prevents drift) | Kelso (can decide) | Open |
| 5 | **skills/ relationship** | Keep both (skills/ = source, plugin/ = artifact) | McCauley/User | Open |
| 6 | **MCP server binary** | PATH assumption (lightweight, fail-fast if missing) | McCauley/User | Open |
| 7 | **Windows-only messaging** | Multi-layered docs (plugin.json + skills + graceful error) | Kelso (can decide) | Open |

**Critical blockers:** Decisions 1, 2, 6 must be resolved + **approval to create published repo** (`sbroenne/mcp-server-excel-plugin`) before Phase 1 can proceed.

---

## Appendix: What We're NOT Building

To avoid scope creep, explicit non-goals:

- ❌ **MCPB bundle** — Separate ecosystem (Claude Desktop), already exists in `mcpb/`
- ❌ **VS Code extension** — Separate distribution method, out of Kelso scope
- ❌ **agentskills.io integration** — Different skill registry, not part of Copilot CLI plugins
- ❌ **Cross-platform support** — Excel is Windows-only, we document this constraint
- ❌ **CLI packaging as skill** — CLI is separate tool, not skill (skills are for AI assistants)
- ❌ **New skill content** — Trejo owns skill CONTENT, Kelso owns PACKAGING
- ❌ **MCP server changes** — Cheritto owns server, Kelso wires config only

---

## Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Plugin spec evolves, breaking our structure | Medium | High | Track GitHub docs, update quarterly |
| MCP server not in PATH | High | Medium | Clear docs, graceful error messages |
| Windows constraint confuses users | Medium | Low | Multi-layered messaging (Decision 7) |
| Marketplace rejects submission | Low | Medium | Start with own-repo distribution (Phase 4) |
| Skills and plugin drift | High | Medium | Build automation (Phase 2) + pre-commit hook |
| Large plugin size (if bundling exe) | Low | Low | Use PATH assumption (Decision 6) |

---

## Timeline Estimate (UPDATED)

Assuming decisions made upfront:

- **Phase 0** (Create published repo): 20 min
- **Phase 1** (MVP): 3 hours
- **Phase 2** (Build automation): 3 hours
- **Phase 3** (Docs): 2 hours
- **Phase 4** (Publish to dedicated repo): 1 hour
- **Phase 5** (Marketplace submission): 2.5 hours (optional, later)
- **Total (MVP → Published):** ~9 hours (1 dev day)
- **Total (MVP → Official Marketplaces):** ~11.5 hours (1.5 dev days)

Post-launch: ~45 min per skill update (rebuild + publish + test).

---

## Next Actions (UPDATED)

1. **User/McCauley:** Review plan + precedent findings, decide on 7 open decisions
2. **Kelso:** Phase 0 — Create published repo `sbroenne/mcp-server-excel-plugin`
3. **Kelso:** Phase 1 — Build MVP plugin in source repo, test locally
4. **Kelso:** Phase 2-3 — Automate build, document installation
5. **Kelso:** Phase 4 — Publish to dedicated repo, test from GitHub
6. **Team:** Validate user experience (register marketplace → install plugin → use)
7. **User:** Decide whether to submit to official marketplaces (Phase 5, optional)

---

**Plan Status:** DRAFT — Awaiting stakeholder decisions + approval to create published repo.

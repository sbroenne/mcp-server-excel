# GitHub Copilot Plugin Distribution

This document outlines how the Excel MCP Server and Excel CLI are distributed as GitHub Copilot CLI plugins through the official marketplace.

## Overview

ExcelMcp is published as **two complementary plugins** in the GitHub Copilot plugin marketplace:

- **`excel-mcp`** — MCP Server with 61 tools across 31 feature areas (387 operations) for conversational AI (Claude Desktop, Copilot chat)
- **`excel-cli`** — CLI-only skill for coding agents (token-efficient, `--help` discoverable)

Both plugins are maintained in a separate published repository and published
from this source repo when their distributed content changes. They are also
listed in [Awesome Copilot](https://github.com/github/awesome-copilot), the
default marketplace in current Copilot clients.

## Distribution Architecture

**Two-Repository Pattern:**
- **This repo** (`sbroenne/mcp-server-excel`) — Source code, release artifacts, plugin templates
- **Published repo** (`sbroenne/mcp-server-excel-plugins`) — GitHub Copilot plugin marketplace artifacts
- **Sync path:** `publish-plugins.yml` builds source-owned templates, validates them, and publishes them to the marketplace

### Why Two Repositories?

- **Plugin marketplace** requires a specific structure with versioned plugin metadata
- **Source repo** focuses on development and component releases
- **Separation of concerns** — release pipeline is independent from plugin packaging

## Plugin Structure (Published Repository)

Each plugin lives in `plugins/` at the published repo:

```
plugins/excel-mcp/
├── plugin.json         # Agent Plugins 1.0 manifest
├── mcp.json            # Portable stdio config that launches npx @latest
├── version.txt         # Published version
├── agents/             # Optional agent definitions
└── skills/             # Optional excel-mcp-report-formatting skill

plugins/excel-cli/
├── plugin.json         # Agent Plugins 1.0 manifest
├── version.txt         # Published version
├── bin/                # Argument-safe PowerShell launcher for npx
└── skills/             # Optional excel-cli-report-formatting skill
```

Agent Plugins discovers skills from the fixed `skills/` directory and MCP servers from root `mcp.json`. The root manifests contain only Agent Plugins 1.0 fields. Skill metadata follows the Agent Skills specification, including name/directory matching and explicit Windows/Excel compatibility.

Each generated plugin receives an exact copy of its canonical skill directory, including every referenced file. This prevents stale published references and preserves skill-specific files such as `references/calculation.md`.

Both plugins use the public npm packages through `npx`; no runtime binaries are
bundled in the plugin package. The MCP plugin's `mcp.json` launches
`npx -y @sbroenne/mcp-server-excel@latest`. The CLI uses
`npx -y @sbroenne/excelcli@latest` and includes `bin\start-cli.ps1` to preserve
quoted JSON arguments on Windows. Node.js 18 or later is required. npm resolves
the `latest` tag and manages package caching, subject to its cache policy; there
is no plugin-owned release downloader or update checker. No global installation
helper, PATH change, or separate global MCP registration is required.
The publish workflow validates this launch-configuration/CLI-wrapper/skill payload
before comparing complete prepared publication output.

## Installation

Users can install either plugin from the default Awesome Copilot marketplace:

```powershell
copilot plugin install excel-mcp@awesome-copilot
copilot plugin install excel-cli@awesome-copilot
```

Alternatively, install from our direct marketplace:

```powershell
# Register the marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install both plugins (or install separately as needed)
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

Choose one marketplace per plugin. Existing direct-marketplace installations
do not need to move; avoid duplicate installations of the same plugin.

### Excel MCP Plugin

Provides the full MCP Server with 61 tools across 31 feature areas (387 operations) for conversational AI:

```powershell
copilot plugin install excel-mcp@awesome-copilot
```

Best for: Claude Desktop, Copilot chat, conversational interfaces.

### Excel CLI Plugin

Provides the argument-safe npx wrapper plus skill guidance for coding agents:

```powershell
copilot plugin install excel-cli@awesome-copilot
```

Best for: CI/CD, scripts, token-efficient coding agents.

## Release Cycle

Plugin publication follows source releases only when complete distributed output
changes beyond known release bookkeeping:

1. **Source release** → `.github/workflows/release.yml` builds all components
2. **Plugin comparison** → `.github/workflows/publish-plugins.yml` publishes real
   changes, or skips commit/push/tag entirely and retains the prior plugin version
3. **Direct marketplace sync** → GitHub Copilot CLI can install current published plugins from our marketplace
4. **Awesome Copilot listing** → An optional upstream update moves its pinned snapshots after review; publication alone does not update those listings

See [Plugin Publishing Workflow Setup](../.github/workflows/docs/publish-plugins-setup.md) for maintainer details.
Product/npm releases continue when plugins are unchanged. Plugin tags are sparse:
not every product release has a matching tag in the published repository.

## Maintenance

Updates to plugins are handled automatically:

1. **Skill updates** → Modify the actual `skills/<name>/SKILL.md` entries or `docs/reference/report-formatting.md`, then run `Build-AgentSkills.ps1 -GenerateOnly`. Other reference documentation is not bundled.
2. **Plugin templates** → Update the canonical `.github/plugins/excel-{mcp,cli}/` sources
3. **Sync to marketplace** → Next release compares complete prepared output,
   including generated references and source-owned root overlays
4. **Awesome Copilot listing** → An optional, disabled-by-default updater maintains
   one upstream PR for actually changed plugin content, using published output
   commits. Root-overlay-only updates do not need a listing PR. Independent
   manual catch-up accepts an existing published tag without republishing.

The published marketplace and the pinned Awesome Copilot listings are separate.
See [Awesome Copilot update setup](../.github/workflows/docs/awesome-copilot-update-setup.md)
for permissions, no-write preview and catch-up.

## Related Documentation

- [Plugin Publishing Workflow](../.github/workflows/docs/publish-plugins-setup.md) — Maintainer guide for plugin release process
- [Release Strategy](RELEASE-STRATEGY.md) — Unified release flow for all components
- [Installation Guide](INSTALLATION.md) — User installation instructions for all clients

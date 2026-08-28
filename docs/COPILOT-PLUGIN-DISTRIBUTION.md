# GitHub Copilot Plugin Distribution

This document outlines how the Excel MCP Server and Excel CLI are distributed as GitHub Copilot CLI plugins through the official marketplace.

## Overview

ExcelMcp is published as **two complementary plugins** in the GitHub Copilot plugin marketplace:

- **`excel-mcp`** — MCP Server with 31 tools (326 operations) for conversational AI (Claude Desktop, Copilot chat)
- **`excel-cli`** — CLI-only skill for coding agents (token-efficient, `--help` discoverable)

Both plugins are maintained in a separate published repository and auto-synced from this source repo.

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
├── mcp.json            # Portable stdio config that launches the bootstrap wrapper
├── version.txt         # Published version
├── bin/                # Portable bootstrap launcher and downloader
├── com.github.copilot/ # Copilot-only global installation helper
├── agents/             # Optional agent definitions
└── skills/             # Behavioral guidance (excel-mcp skill)

plugins/excel-cli/
├── plugin.json         # Agent Plugins 1.0 manifest
├── version.txt         # Published version
├── bin/                # Portable bootstrap launcher and downloader
├── com.github.copilot/ # Copilot-only global installation helper
└── skills/             # Behavioral guidance (excel-cli skill)
```

Agent Plugins discovers skills from the fixed `skills/` directory and MCP servers from root `mcp.json`. The root manifests contain only Agent Plugins 1.0 fields; any future Copilot-only files must live under `com.github.copilot/`. Skill metadata follows the Agent Skills specification, including name/directory matching and explicit Windows/Excel compatibility.

Each generated plugin receives an exact copy of its canonical skill directory, including every referenced file. This prevents stale published references and preserves skill-specific files such as `references/calculation.md`.

Both plugins publish **wrapper/bootstrap assets only** — no runtime binaries are bundled in the plugin package. On first use, each plugin downloads and caches the newest self-contained Windows runtime (`mcp-excel.exe` or `excelcli.exe`) from the main repo's GitHub Releases feed. The bootstrap reads the exact release's `SHA256SUMS` asset and verifies the selected ZIP before extraction; cached ZIPs are verified again before reuse. Missing, malformed, unmatched, or incorrect checksum data stops installation. Agent Plugins hosts provide `PLUGIN_DATA`; the bootstrap stores persistent runtime state under `PLUGIN_DATA\runtime`, checks release freshness once per Copilot session, and then reuses the verified runtime. Standalone shims fall back to `~\.copilot\plugin-runtime\mcp-server-excel\<plugin>` and check for updates at most once every 24 hours. The publish workflow validates this wrapper/bootstrap-only payload before syncing to the marketplace repo.

## Installation

Users install the two plugins directly from the GitHub Copilot CLI marketplace:

```powershell
# Register the marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install both plugins (or install separately as needed)
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

### Excel MCP Plugin

Provides the full MCP Server with 31 tools (326 operations) for conversational AI:

```powershell
copilot plugin install excel-mcp@mcp-server-excel-plugins
```

Best for: Claude Desktop, Copilot chat, conversational interfaces.

### Excel CLI Plugin

Provides the CLI bootstrap wrapper plus skill guidance for coding agents:

```powershell
copilot plugin install excel-cli@mcp-server-excel-plugins
pwsh -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\com.github.copilot\bin\install-global.ps1"
```

Best for: CI/CD, scripts, token-efficient coding agents.

## Release Cycle

Both plugins are republished automatically after each source repo release:

1. **Source release** → `.github/workflows/release.yml` builds all components
2. **Plugin publish** → `.github/workflows/publish-plugins.yml` syncs to marketplace repo
3. **Marketplace sync** → GitHub Copilot CLI discovers both plugins

See [Plugin Publishing Workflow Setup](../.github/workflows/docs/publish-plugins-setup.md) for maintainer details.

## Maintenance

Updates to plugins are handled automatically:

1. **Skill updates** → Modify `skills/excel-mcp/` or `skills/excel-cli/` in this repo
2. **Plugin templates** → Update the canonical `.github/plugins/excel-{mcp,cli}/` sources
3. **Sync to marketplace** → Next release runs `publish-plugins.yml` to update both plugins
4. **No awesome-copilot PR needed** — Plugins are fetched from the published marketplace repo

This approach keeps plugin distribution simple — users always see the latest version from the marketplace, and maintainers only need to manage one source repo and one published repo.

## Related Documentation

- [Plugin Publishing Workflow](../.github/workflows/docs/publish-plugins-setup.md) — Maintainer guide for plugin release process
- [Release Strategy](RELEASE-STRATEGY.md) — Unified release flow for all components
- [Installation Guide](INSTALLATION.md) — User installation instructions for all clients

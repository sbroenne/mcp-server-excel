# GitHub Copilot Plugin Distribution

This document outlines how the Excel MCP Server and Excel CLI are distributed as GitHub Copilot CLI plugins through the official marketplace.

## Overview

ExcelMcp is published as **two complementary plugins** in the GitHub Copilot plugin marketplace:

- **`excel-mcp`** — MCP Server with 25 tools (230 operations) for conversational AI (Claude Desktop, Copilot chat)
- **`excel-cli`** — CLI-only skill for coding agents (token-efficient, `--help` discoverable)

Both plugins are maintained in a separate published repository and auto-synced from this source repo.

## Distribution Architecture

**Two-Repository Pattern:**
- **This repo** (`sbroenne/mcp-server-excel`) — Source code, release artifacts, plugin templates
- **Published repo** (`sbroenne/mcp-server-excel-plugins`) — GitHub Copilot plugin marketplace artifacts
- **Sync path:** `publish-plugins.yml` workflow copies templates, applies overlays, and publishes to marketplace

### Why Two Repositories?

- **Plugin marketplace** requires a specific structure with versioned plugin metadata
- **Source repo** focuses on development and component releases
- **Separation of concerns** — release pipeline is independent from plugin packaging

## Plugin Structure (Published Repository)

Each plugin lives in `plugins/` at the published repo:

```
plugins/excel-mcp/
├── plugin.json         # MCP Server + skill metadata
├── .mcp.json           # MCP Server configuration
├── version.txt         # Published version
├── bin/                # MCP Server executable
├── agents/             # Optional agent definitions
└── skills/             # Behavioral guidance (excel-mcp skill)

plugins/excel-cli/
├── plugin.json         # CLI-only metadata
├── version.txt         # Published version
├── bin/                # CLI executable (excelcli.exe)
└── skills/             # Behavioral guidance (excel-cli skill)
```

The skills reference is shared from this source repo (`skills/shared/*.md`).

## Installation

Users install the two plugins directly from the GitHub Copilot CLI marketplace:

```powershell
# Register the marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install both plugins (or install separately as needed)
copilot plugin install excel-mcp@sbroenne/mcp-server-excel-plugins
copilot plugin install excel-cli@sbroenne/mcp-server-excel-plugins
```

### Excel MCP Plugin

Provides the full MCP Server with 25 tools (230 operations) for conversational AI:

```powershell
copilot plugin install excel-mcp@sbroenne/mcp-server-excel-plugins
```

Best for: Claude Desktop, Copilot chat, conversational interfaces.

### Excel CLI Plugin

Provides the CLI tool bundled with skill guidance for coding agents:

```powershell
copilot plugin install excel-cli@sbroenne/mcp-server-excel-plugins
pwsh -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
```

Best for: CI/CD, scripts, token-efficient coding agents.

## Release Cycle

Both plugins are republished automatically after each source repo release:

1. **Source release** → `.github/workflows/release.yml` builds all components
2. **Plugin publish** → `.github/workflows/publish-plugins.yml` syncs to marketplace repo
3. **Marketplace sync** → GitHub Copilot CLI discovers both plugins

See [Plugin Publishing Workflow Setup](../workflows/docs/publish-plugins-setup.md) for maintainer details.

## Maintenance

Updates to plugins are handled automatically:

1. **Skill updates** → Modify `skills/excel-mcp/` or `skills/excel-cli/` in this repo
2. **Plugin templates** → Update `.github/plugins/excel-{mcp,cli}/` overlays
3. **Sync to marketplace** → Next release runs `publish-plugins.yml` to update both plugins
4. **No awesome-copilot PR needed** — Plugins are fetched from the published marketplace repo

This approach keeps plugin distribution simple — users always see the latest version from the marketplace, and maintainers only need to manage one source repo and one published repo.

## Related Documentation

- [Plugin Publishing Workflow](../workflows/docs/publish-plugins-setup.md) — Maintainer guide for plugin release process
- [Release Strategy](RELEASE-STRATEGY.md) — Unified release flow for all components
- [Installation Guide](INSTALLATION.md) — User installation instructions for all clients

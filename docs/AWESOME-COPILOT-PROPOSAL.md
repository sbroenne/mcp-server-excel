# Awesome Copilot Contribution Proposal

This document outlines the plan for contributing the Excel MCP Server as a plugin to [github/awesome-copilot](https://github.com/github/awesome-copilot).

## Overview

The Excel MCP Server provides 25 specialized tools with 225+ operations for comprehensive Excel automation through AI assistants. Contributing it as an awesome-copilot plugin makes it discoverable to the broader GitHub Copilot community.

## Contribution Type: External Plugin

We contribute as an **external plugin** — the plugin definition lives in this repository and is referenced from awesome-copilot's `plugins/external.json`.

### Why External Plugin?

- Plugin source files stay in this repo alongside the MCP Server code
- Updates to the agent, skill, or plugin metadata don't require PRs to awesome-copilot
- The skill references and agent file remain synchronized with the actual tool implementation
- Version management stays under our control

## Plugin Structure (This Repository)

The plugin is defined at `.github/plugins/excel-automation/`:

```
.github/plugins/excel-automation/
├── plugin.json                              # Plugin metadata (Claude Code spec)
├── README.md                                # Plugin documentation
└── agents/
    └── excel-automation.agent.md            # Excel automation agent definition
```

The plugin references the existing skill at `skills/excel-mcp/` for detailed tool guidance.

## PR to awesome-copilot

### Step 1: Fork and Branch

```bash
# Fork github/awesome-copilot
# Clone your fork
git clone https://github.com/<your-username>/awesome-copilot.git
cd awesome-copilot
git checkout -b add-excel-automation-plugin
```

### Step 2: Add External Plugin Entry

Edit `plugins/external.json` and add the following entry to the array:

```json
{
  "name": "excel-automation",
  "description": "Automate Microsoft Excel on Windows through natural language. 25 tools with 225+ operations covering Power Query, DAX, PivotTables, Charts, VBA, Tables, Ranges, Slicers, and more via Excel's native COM API.",
  "version": "1.0.0",
  "author": {
    "name": "Stefan Broenner",
    "url": "https://github.com/sbroenne"
  },
  "homepage": "https://github.com/sbroenne/mcp-server-excel",
  "keywords": [
    "excel",
    "spreadsheet",
    "automation",
    "mcp",
    "power-query",
    "dax",
    "pivottable",
    "charts",
    "vba",
    "com-interop",
    "windows"
  ],
  "license": "MIT",
  "repository": "https://github.com/sbroenne/mcp-server-excel",
  "source": {
    "source": "github",
    "repo": "sbroenne/mcp-server-excel",
    "path": ".github/plugins/excel-automation"
  }
}
```

### Step 3: Build and Validate

```bash
npm install
npm run build
npm run plugin:validate
```

### Step 4: Submit PR

Submit a PR to the `staged` branch (not `main`) of `github/awesome-copilot` with:

- **Title**: `Add Excel Automation plugin (MCP Server for Excel)`
- **Description**: See [PR Description Template](#pr-description-template) below

> **Important**: All PRs to awesome-copilot must target the `staged` branch.

## PR Description Template

```markdown
## Add Excel Automation Plugin

### What This Plugin Provides

**Excel Automation** — Automate Microsoft Excel on Windows through natural language using the [Excel MCP Server](https://github.com/sbroenne/mcp-server-excel).

**25 tools with 225+ operations** covering:
- Power Query (M code import, evaluate, refresh)
- Data Model / DAX (measures, relationships)
- PivotTables (create, fields, calculated items)
- Charts (create, configure, series, trendlines)
- Tables (create, filter, sort, structured references)
- Ranges (values, formulas, formatting, validation)
- VBA (modules, macro execution)
- Slicers, Named Ranges, Connections, and more

### Plugin Type
External plugin — source hosted at [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel).

### What's Included
- **Agent**: `excel-automation` — Excel automation expert agent
- **Skill**: `excel-mcp` — Comprehensive MCP Server skill with workflow guidance

### Prerequisites
- Windows with Microsoft Excel 2016+
- .NET 10 SDK
- Install: `dotnet tool install --global Sbroenne.ExcelMcp.McpServer`

### Testing
- Plugin structure validated with `npm run plugin:validate`
- MCP Server is published on NuGet and actively maintained
- Tool behavior validated with real LLM workflows using [pytest-aitest](https://github.com/sbroenne/pytest-aitest)

### Related Links
- Repository: https://github.com/sbroenne/mcp-server-excel
- NuGet: https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer
- VS Code Extension: https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp
```

## Contributor Recognition

After the PR is merged, request contributor recognition by commenting:

```markdown
@all-contributors add @sbroenne for plugins, agents
```

## Maintenance

When the plugin needs updates:

1. Update files in `.github/plugins/excel-automation/` in this repository
2. If only content changes (agent/skill updates), no PR to awesome-copilot needed — external plugins are fetched from source
3. If metadata changes (name, description, keywords), update the `plugins/external.json` entry in awesome-copilot

## Alternative: Standalone Agent Contribution

If the external plugin approach is not accepted, we can alternatively contribute just the agent file directly to awesome-copilot:

1. Copy `excel-automation.agent.md` to `agents/` in awesome-copilot
2. Add the MCP server reference in the agents README table
3. Submit PR targeting `staged` branch

This is simpler but doesn't include the skill, and the agent file would need to be maintained in both repositories.

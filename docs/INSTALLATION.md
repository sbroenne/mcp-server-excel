# Installation Guide - ExcelMcp

Install ExcelMcp to automate installed Microsoft Excel on Windows from an AI
assistant or the command line. It ships two **equal entry points** — the
**MCP Server** for MCP clients and the **CLI** (`excelcli`) for coding agents
and scripts. Choose the guide for your workflow; the entry points are
independent:

| Guide | Best For |
|-------|----------|
| 📖 **[Installing the MCP Server](INSTALLATION-MCP-SERVER.md)** | AI assistants — GitHub Copilot, Claude Desktop, Cursor, Windsurf, and any other MCP client |
| 📖 **[Installing the CLI](INSTALLATION-CLI.md)** | Scripting, RPA, CI/CD pipelines, and coding agents that prefer a token-efficient single tool |

Both require **Windows OS**, **Microsoft Excel 2016+**, and an **interactive
desktop**. npm and standalone executables need no separate .NET runtime;
NuGet tools require .NET 10. Manual npm use requires Node.js 18+.

| Where you work | Recommended installation |
|---|---|
| VS Code with GitHub Copilot | VS Code extension; bundles the server and its skill |
| Claude Desktop | [Claude Desktop setup guide](guides/CLAUDE-DESKTOP.md); MCPB configures direct npx with `@latest` (Node.js/npm required) |
| Another MCP client | npm through `npx -y @sbroenne/mcp-server-excel@latest` |
| Coding agents and scripts | `npx -y @sbroenne/excelcli@latest`, or global npm for a command on PATH |
| No npm downloads desired | Standalone ZIP; replace the executable manually for updates |

`@latest` is resolved when launching, subject to normal npm caching. It does not
upgrade a running server or CLI background service. The guides below explain
safe restarts and updates for each method.

> **Tip:** The **VS Code Extension** bundles the MCP Server only (install the CLI separately if you need it for scripting). The **GitHub Copilot plugins** are separate — install `excel-mcp` and/or `excel-cli` depending on which entry point you need — see the MCP Server guide's Quick Start for the one-click paths.

The Copilot plugins are also listed in
[Awesome Copilot](https://github.com/github/awesome-copilot), the default
marketplace in current Copilot clients. Each installation guide includes that
install route and our direct marketplace alternative.

---

## Agent Skills Installation (Cross-Platform)

**Best for:** Optional presentation guidance for requested Excel reports.

The VS Code extension registers `excel-mcp-report-formatting`. Plugins remain
`excel-mcp` and `excel-cli`; their contained skills have narrower names and scope.
After the release containing this change is published, install directly with:

```powershell
# CLI report-formatting skill
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting

# MCP report-formatting skill
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp-report-formatting

# Interactive install - select one or both formatting skills
npx skills add sbroenne/mcp-server-excel-plugins

# Install for specific agents
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting -a cursor
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp-report-formatting -a claude-code

# Install both skills
npx skills add sbroenne/mcp-server-excel-plugins --skill '*'

# Install globally (user-wide)
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli-report-formatting --global
```

**Supports 43+ agents** including claude-code, github-copilot, cursor, windsurf, gemini-cli, codex, goose, cline, continue, replit, and more.

**Manual Installation:**

Existing broad `excel-cli` and `excel-mcp` standalone skills remain installed
until removed with the client's skill manager. Keep plugin/server configuration.
Use complete prepared packages, not source entries without their references.

1. Download `excel-skills-v{version}.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. The package contains both skills:
   - `skills/excel-cli-report-formatting/` - presentation through `excelcli`
   - `skills/excel-mcp-report-formatting/` - presentation through MCP tools
3. Extract the skill(s) you need to your AI assistant's skills directory:
   - Copilot: `~/.copilot/skills/<skill-name>/`
   - Claude Code: `.claude/skills/<skill-name>/`
   - Cursor: `.cursor/skills/<skill-name>/`

**See:** [Agent Skills Documentation](../docs/AGENT-SKILLS.md)

---

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

**Happy automating! 🚀**

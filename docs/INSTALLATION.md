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

## Optional Report-Formatting Guidance

**Best for:** Optional presentation guidance for requested Excel reports.

The `excel-mcp` and `excel-cli` plugins include report-formatting guidance.
The VS Code extension registers `excel-mcp-report-formatting`. No separate
skill installation or download is needed.

**See:** [Agent Skills Documentation](../docs/AGENT-SKILLS.md)

---

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

**Happy automating! 🚀**

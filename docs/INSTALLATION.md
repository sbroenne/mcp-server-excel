# Installation Guide - ExcelMcp

ExcelMcp lets an AI assistant or coding agent work with Microsoft Excel
installed on Windows. There are two ways to use it, and you can install either
or both:

- **MCP Server**: for AI assistants that support MCP, such as GitHub Copilot,
  Claude Desktop, Cursor, and Windsurf.
- **CLI** (`excelcli`): the same features as terminal commands. Coding agents
  can use either; the CLI usually uses fewer AI tokens.

## What you need

- Windows
- Desktop Microsoft Excel 2016 or later
- A signed-in Windows desktop session, because Excel runs as a desktop app
- Node.js 18 or later for the GitHub Copilot plugins, the Claude Desktop
  extension, or installation through npm

## Pick your setup

| Where you work | What to install | Guide |
|---|---|---|
| VS Code with GitHub Copilot | The Excel MCP Server VS Code extension | [VS Code extension](INSTALLATION-MCP-SERVER.md#vs-code-extension-easiest-one-click-setup) |
| GitHub Copilot plugins | The `excel-mcp` plugin, the `excel-cli` plugin, or both; listed in [Awesome Copilot](https://github.com/github/awesome-copilot) | [MCP plugin](INSTALLATION-MCP-SERVER.md#github-copilot-plugin), [CLI plugin](INSTALLATION-CLI.md#github-copilot-plugin) |
| Claude Desktop | The one-click Claude Desktop extension | [Claude Desktop setup](guides/CLAUDE-DESKTOP.md) |
| Another MCP client | The MCP Server through npm | [Manual MCP setup](INSTALLATION-MCP-SERVER.md#manual-mcp-setup-all-mcp-clients) |
| Coding agents (fewer AI tokens) | The CLI through npm | [CLI with npm](INSTALLATION-CLI.md#npm-primary) |
| A computer without npm | The standalone MCP Server or CLI download | [MCP Server download](INSTALLATION-MCP-SERVER.md#option-b-standalone-executable), [CLI download](INSTALLATION-CLI.md#standalone-executable-also-primary) |

The VS Code extension sets up the MCP Server only. If your agent also runs
terminal commands, install the CLI too.

In GitHub Copilot, install the `excel-mcp` plugin, the `excel-cli` plugin, or
both.

The plugins and the VS Code extension already include guidance for formatting
Excel reports, so you do not need to install anything else.

Each guide also covers updating, troubleshooting, and uninstalling.

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

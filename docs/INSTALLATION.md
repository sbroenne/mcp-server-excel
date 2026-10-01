# Installation Guide - ExcelMcp

ExcelMcp ships two **equal entry points** — the **MCP Server** for AI assistants
and the **CLI** for scripting, RPA, and coding agents on an interactive desktop
Excel host. Headless CI is unsupported. Pick the guide that matches how you'll
use it (or read both, they're independent):

| Guide | Best For |
|-------|----------|
| 📖 **[Installing the MCP Server](INSTALLATION-MCP-SERVER.md)** | AI assistants — GitHub Copilot, Claude Desktop, Cursor, Windsurf, and any other MCP client |
| 📖 **[Installing the CLI](INSTALLATION-CLI.md)** | Scripting, RPA, and coding agents on a desktop Excel host |

Both entry points support **Windows with Microsoft Excel 2016+** and
**Apple Silicon macOS with Excel for Mac 16.112+**. Windows provides the complete
operation set; macOS support is **experimental beta** with a
[capability-gated subset and explicit limitations](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta).
The standalone distributions do not require a .NET runtime.

> **macOS beta exclusions:** Power Query, VBA, Data Model/DAX/OLAP, Tables,
> PivotTables, charts, slicers, connections, QueryTables, XML Maps, screenshots,
> advanced visual formatting, and Python result reads. Test on workbook copies;
> installing a different entry point or the optional bridge does not enable
> these features.

Windows and Apple Silicon macOS use separate native archives; one executable
file cannot be shared across PE/Windows and Mach-O/macOS. Intel macOS is
unsupported and fails closed rather than selecting the ARM64 runtime. See
[macOS distribution readiness](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/MACOS-DISTRIBUTION.md)
for package inspection, signing, and notarization details.

The [optional macOS Office.js bridge](MACOS-OFFICEJS.md) is a development-stage
capability foundation with separate explicit install, activation, health,
upgrade, and removal steps. It is not needed for the base macOS operation set
and does not currently enable tables, charts, PivotTables, or formatting.

> **Tip:** The **VS Code Extension** bundles the MCP Server only (install the CLI separately if you need it for scripting). The **GitHub Copilot plugins** are separate — install `excel-mcp` and/or `excel-cli` depending on which entry point you need — see the MCP Server guide's Quick Start for the one-click paths.

---

## Agent Skills Installation (Cross-Platform)

**Best for:** Adding AI guidance to coding agents (Copilot, Cursor, Windsurf, Claude Code, Gemini, Codex, etc.)

The VS Code extension auto-installs the `excel-mcp` skill only. Plugins and skills are different things: plugins are packaged surface integrations, while skills are reusable AI guidance. For the `excel-cli` skill, or for environments where you want skills directly, use the commands below:

```powershell
# CLI skill (for coding agents - token-efficient workflows)
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli

# MCP skill (for conversational AI - rich tool schemas)
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp

# Interactive install - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel-plugins

# Install for specific agents
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli -a cursor
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-mcp -a claude-code

# Install both skills
npx skills add sbroenne/mcp-server-excel-plugins --skill '*'

# Install globally (user-wide)
npx skills add sbroenne/mcp-server-excel-plugins --skill excel-cli --global
```

**Supports 43+ agents** including claude-code, github-copilot, cursor, windsurf, gemini-cli, codex, goose, cline, continue, replit, and more.

**Manual Installation:**

Existing skills remain installed. The old source-repository installation command
does not redirect; use `sbroenne/mcp-server-excel-plugins` for future installs and updates.

1. Download `excel-skills-v{version}.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. The package contains both skills:
   - `skills/excel-cli/` - for coding agents (Copilot, Cursor, Windsurf)
   - `skills/excel-mcp/` - for conversational AI (Claude Desktop, VS Code Chat)
3. Extract the skill(s) you need to your AI assistant's skills directory:
   - Copilot: `~/.copilot/skills/excel-cli/` or `~/.copilot/skills/excel-mcp/`
   - Claude Code: `.claude/skills/excel-cli/` or `.claude/skills/excel-mcp/`
   - Cursor: `.cursor/skills/excel-cli/` or `.cursor/skills/excel-mcp/`

**See:** [Agent Skills Documentation](../skills/README.md)

---

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

**Happy automating! 🚀**

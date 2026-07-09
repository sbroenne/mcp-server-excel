# Installation Guide - ExcelMcp

<!--start-->
ExcelMcp ships two **equal entry points** — the **MCP Server** for AI assistants and the **CLI** for scripting, RPA, and CI/CD. Pick the guide that matches how you'll use it (or read both, they're independent):

| Guide | Best For |
|-------|----------|
| 📖 **[Installing the MCP Server](https://excelmcpserver.dev/installation-mcp-server/)** | AI assistants — GitHub Copilot, Claude Desktop, Cursor, Windsurf, and any other MCP client |
| 📖 **[Installing the CLI](https://excelmcpserver.dev/installation-cli/)** | Scripting, RPA, CI/CD pipelines, and coding agents that prefer a token-efficient single tool |

Both require **Windows OS** and **Microsoft Excel 2016+** — no .NET runtime needed for the standalone exe distributions.

> **Tip:** The **VS Code Extension** bundles the MCP Server only (install the CLI separately if you need it for scripting). The **GitHub Copilot plugins** are separate — install `excel-mcp` and/or `excel-cli` depending on which entry point you need — see the MCP Server guide's Quick Start for the one-click paths.

---

## Agent Skills Installation (Cross-Platform)

**Best for:** Adding AI guidance to coding agents (Copilot, Cursor, Windsurf, Claude Code, Gemini, Codex, etc.)

The VS Code extension auto-installs the `excel-mcp` skill only. Plugins and skills are different things: plugins are packaged surface integrations, while skills are reusable AI guidance. For the `excel-cli` skill, or for environments where you want skills directly, use the commands below:

```powershell
# CLI skill (for coding agents - token-efficient workflows)
npx skills add sbroenne/mcp-server-excel --skill excel-cli

# MCP skill (for conversational AI - rich tool schemas)
npx skills add sbroenne/mcp-server-excel --skill excel-mcp

# Interactive install - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel

# Install for specific agents
npx skills add sbroenne/mcp-server-excel --skill excel-cli -a cursor
npx skills add sbroenne/mcp-server-excel --skill excel-mcp -a claude-code

# Install both skills
npx skills add sbroenne/mcp-server-excel --skill '*'

# Install globally (user-wide)
npx skills add sbroenne/mcp-server-excel --skill excel-cli --global
```

**Supports 43+ agents** including claude-code, github-copilot, cursor, windsurf, gemini-cli, codex, goose, cline, continue, replit, and more.

**Manual Installation:**
1. Download `excel-skills-v{version}.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. The package contains both skills:
   - `skills/excel-cli/` - for coding agents (Copilot, Cursor, Windsurf)
   - `skills/excel-mcp/` - for conversational AI (Claude Desktop, VS Code Chat)
3. Extract the skill(s) you need to your AI assistant's skills directory:
   - Copilot: `~/.copilot/skills/excel-cli/` or `~/.copilot/skills/excel-mcp/`
   - Claude Code: `.claude/skills/excel-cli/` or `.claude/skills/excel-mcp/`
   - Cursor: `.cursor/skills/excel-cli/` or `.cursor/skills/excel-mcp/`

**See:** [Agent Skills Documentation](https://excelmcpserver.dev/skills/)

<!--end-->

---

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](https://excelmcpserver.dev/contributing/)

**Happy automating! 🚀**

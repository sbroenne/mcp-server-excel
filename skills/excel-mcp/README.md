# Excel MCP Server Skill

Agent Skill for AI assistants using the Excel MCP Server via the Model Context Protocol.

## Best For

- **Conversational AI** (Claude Desktop, VS Code Chat)
- Exploratory automation with iterative reasoning
- Self-healing workflows needing rich introspection
- Long-running autonomous tasks with continuous context

## Installation

### GitHub Copilot

The [Excel MCP Server VS Code extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) installs this skill automatically to `~/.copilot/skills/excel-mcp/`.

Enable skills in VS Code settings:
```json
{
  "chat.useAgentSkills": true
}
```

### Other Platforms

Extract to your AI assistant's skills directory:

| Platform | Location |
|----------|----------|
| **Claude Code** | `.claude/skills/excel-mcp/` |
| **Cursor** | `.cursor/skills/excel-mcp/` |
| **Windsurf** | `.windsurf/skills/excel-mcp/` |
| **Gemini CLI** | `.gemini/skills/excel-mcp/` |
| **Codex** | `.codex/skills/excel-mcp/` |
| **Goose** | `.goose/skills/excel-mcp/` |
| **And 36+ more** | Via `npx skills` |

Or use npx:
```powershell
# Interactive - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel

# Or specify directly
npx skills add sbroenne/mcp-server-excel --skill excel-mcp
```

## Contents

```
excel-mcp/
├── SKILL.md           # Main skill definition with MCP tool guidance
├── README.md          # This file
├── VERSION             # Published plugin version
└── references/        # Detailed domain-specific guidance
    └── *.md
```

Distributable packages add a `VERSION` file during the build. The canonical skill
source intentionally has no version metadata so it cannot become a stale build input.

## MCP Server Setup

The skill works with the Excel MCP Server. See [Installation Guide](https://excelmcpserver.dev/installation/) for setup instructions.

## Related

- [Excel CLI Plugin](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli) - For coding agents preferring CLI tools
- [Documentation](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)

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
| **Goose** | `.goose/skills/excel-mcp/` |

Or use npx:
```bash
npx add-skill sbroenne/mcp-server-excel --skill excel-mcp -a claude-code
```

## Contents

```
excel-mcp/
├── SKILL.md           # Main skill definition with MCP tool guidance
├── VERSION            # Version tracking
├── README.md          # This file
└── references/        # Detailed domain-specific guidance
    ├── behavioral-rules.md
    ├── anti-patterns.md
    ├── workflows.md
    ├── excel_range.md
    ├── excel_table.md
    ├── excel_worksheet.md
    ├── excel_chart.md
    ├── excel_slicer.md
    ├── excel_powerquery.md
    ├── excel_datamodel.md
    ├── excel_conditionalformat.md
    └── claude-desktop.md
```

## MCP Server Setup

The skill works with the Excel MCP Server. See [Installation Guide](https://excelmcpserver.dev/installation/) for setup instructions.

## Related

- [Excel CLI Skill](https://github.com/sbroenne/mcp-server-excel/releases) - For coding agents preferring CLI tools
- [Documentation](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)

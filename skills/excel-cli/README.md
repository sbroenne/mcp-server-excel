# Excel CLI Skill

Agent Skill for AI coding assistants using the Excel CLI tool (`excelcli`).

## Best For

- **Coding agents** (GitHub Copilot, Cursor, Windsurf)
- Token-efficient workflows (no large tool schemas)
- Discoverable via `excelcli --help`
- Scriptable in PowerShell pipelines, CI/CD, batch processing
- Quiet mode (`-q`) outputs clean JSON only

## Why CLI Over MCP?

Modern coding agents increasingly favor CLI-based workflows:

```powershell
# Token-efficient: No schema overhead
excelcli -q session open C:\Data\Report.xlsx
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values-json '[["Hello"]]'
excelcli -q session close --session 1 --save
```

## Installation

### GitHub Copilot

The [Excel MCP Server VS Code extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) installs this skill automatically to `~/.copilot/skills/excel-cli/`.

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
| **Claude Code** | `.claude/skills/excel-cli/` |
| **Cursor** | `.cursor/skills/excel-cli/` |
| **Windsurf** | `.windsurf/skills/excel-cli/` |
| **Gemini CLI** | `.gemini/skills/excel-cli/` |
| **Goose** | `.goose/skills/excel-cli/` |

Or use npx:
```bash
npx add-skill sbroenne/mcp-server-excel --skill excel-cli -a cursor
```

## Contents

```
excel-cli/
├── SKILL.md           # Main skill definition with CLI command guidance
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
    └── excel_conditionalformat.md
```

## CLI Installation

Install the CLI tool via NuGet:
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

Verify installation:
```powershell
excelcli --version
excelcli --help
```

## Related

- [Excel MCP Skill](https://github.com/sbroenne/mcp-server-excel/releases) - For conversational AI (Claude Desktop, VS Code Chat)
- [Documentation](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)

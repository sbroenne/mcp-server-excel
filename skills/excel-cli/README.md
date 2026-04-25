# Excel CLI Skill

Agent Skill for AI coding assistants using the Excel CLI tool (`excelcli`).

## Best For

- **Coding agents** (GitHub Copilot, Cursor, Windsurf, Codex, Gemini CLI, and 38+ more)
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
| **Codex** | `.codex/skills/excel-cli/` |
| **And 36+ more** | Via `npx skills` |
| **Goose** | `.goose/skills/excel-cli/` |

Or use npx:
```bash
# Interactive - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel

# Or specify directly
npx skills add sbroenne/mcp-server-excel --skill excel-cli
```

## Contents

```
excel-cli/
├── SKILL.md           # Main skill definition with CLI command guidance
├── README.md          # This file
└── references/        # Detailed domain-specific guidance
    ├── anti-patterns.md
    ├── behavioral-rules.md
    ├── chart.md
    ├── conditionalformat.md
    ├── dashboard.md
    ├── datamodel.md
    ├── dmv-reference.md
    ├── excel_agent_mode.md
    ├── gotchas.md
    ├── m-code-syntax.md
    ├── pivottable.md
    ├── powerquery.md
    ├── range.md
    ├── screenshot.md
    ├── slicer.md
    ├── table.md
    ├── window.md
    └── worksheet.md
```

## CLI Tool Installation

The **GitHub Copilot `excel-cli` plugin** now bundles the self-contained `excelcli.exe` deliverable (no .NET runtime required for that install path).

### Via GitHub Copilot Plugin

If you install `excel-cli` through the GitHub Copilot plugin marketplace, the plugin ships the actual CLI in its `bin\` folder. Run the bundled helper once to expose it on PATH:

```powershell
pwsh -ExecutionPolicy Bypass -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
```

### Via Skill Package

Plain skill-only installs still need `excelcli` available separately on PATH (for example via the standalone ZIP or the NuGet tool below).

If you reinstall the plugin to a different location, re-run `install-global.ps1` so the shim points at the current bundled binary.

### Manual Download (Standalone)

For other environments, download the standalone CLI:
```powershell
# Download from releases
$url = "https://github.com/sbroenne/mcp-server-excel/releases/latest/download/ExcelMcp-CLI-latest-windows.zip"
Invoke-WebRequest -Uri $url -OutFile ExcelMcp-CLI.zip
Expand-Archive -Path ExcelMcp-CLI.zip -DestinationPath $env:ProgramFiles\ExcelMcp
```

### Via NuGet Package Manager (Secondary)

Requires .NET 10 Runtime or SDK:
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

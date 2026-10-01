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
# Discover a session rather than inventing an ID.
excelcli -q session list
excelcli -q range get-values --session $sessionId --sheet Sheet1 --range A1
```

Here `$sessionId` is the returned ID for the intended workbook. For writes and
batch jobs, use the failure-aware lifecycle in [SKILL.md](SKILL.md).

## Installation

### GitHub Copilot

The VS Code extension bundles only the MCP Server skill. Install this CLI skill
separately with `npx skills`, as shown below, and run the CLI through `npx`.

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
```powershell
# Interactive - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel-plugins

# Or specify directly
npx skills add https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli/skills/excel-cli
```

## Contents

```
excel-cli/
├── SKILL.md           # Main skill definition with CLI command guidance
├── README.md          # This file
├── VERSION            # Published plugin version
└── references/        # CLI command reference and workflow guidance
    └── *.md
```

## CLI Tool Installation

The **GitHub Copilot `excel-cli` plugin** installs the skill plus an npx-first
wrapper for the public `@sbroenne/excelcli` package.

### Via GitHub Copilot Plugin

Use the public npm package directly:

```powershell
npx -y @sbroenne/excelcli@latest --help
```

The plugin also includes `bin\start-cli.ps1`, which runs npx and preserves
embedded quotes in JSON arguments from Windows PowerShell. No global helper
or PATH change is required.

### Via Skill Package

Plain skill-only installs can use `npx -y @sbroenne/excelcli@latest` without a
separate CLI installation. Replace `excelcli` in the examples with that command
unless you have installed a standalone CLI on PATH.

### Manual Download (Standalone)

For other environments, download `ExcelMcp-CLI-{version}-windows.zip` from the
[latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest),
extract it to a permanent directory, and add that directory to PATH.

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

- [Excel MCP Skill](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-mcp/skills/excel-mcp) - For conversational AI (Claude Desktop, VS Code Chat)
- [Documentation](https://excelmcpserver.dev/)
- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)

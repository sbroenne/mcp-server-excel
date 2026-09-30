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
separately with `npx skills`, as shown below, after installing `excelcli`.

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

Use `npx -y @sbroenne/excelcli@latest` directly, or run the optional global shim
installer from the installed plugin folder to make `excelcli` available on PATH:

```powershell
pwsh -ExecutionPolicy Bypass -File `
  "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\com.github.copilot\bin\install-global.ps1"
```

The global shim uses npx and preserves embedded quotes in JSON arguments from
Windows PowerShell.

### Via Skill Package

Plain skill-only installs still need `excelcli` available separately on PATH (for example via the standalone ZIP or the NuGet tool below).

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

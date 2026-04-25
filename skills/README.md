# Excel MCP Server - Agent Skills

Two skill packages bundled in GitHub Copilot plugins and distributed separately:

| Skill | Component | Distribution | Best For |
|-------|-----------|--------------|----------|
| **[excel-cli](excel-cli/SKILL.md)** | CLI Tool (`excelcli.exe`) | Copilot plugin `excel-cli`, npx, NuGet | Coding agents - token-efficient, `--help` discoverable |
| **[excel-mcp](excel-mcp/SKILL.md)** | MCP Server (`mcp-excel.exe`) | Copilot plugin `excel-mcp`, VS Code extension, MCPB | Conversational AI - rich tool schemas |

**Shared guidance:** `skills/shared/*.md` — source of truth for both skills (auto-copied to each skill's `references/` folder)

## Installation

**GitHub Copilot Plugins (Recommended):**
```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@sbroenne/mcp-server-excel-plugins
copilot plugin install excel-cli@sbroenne/mcp-server-excel-plugins
```

**Manual skill extraction:**
```bash
# Via npx (interactive — select excel-cli, excel-mcp, or both)
npx skills add sbroenne/mcp-server-excel

# Or specify directly
npx skills add sbroenne/mcp-server-excel --skill excel-cli
npx skills add sbroenne/mcp-server-excel --skill excel-mcp
```

**Via VS Code Extension (auto-installs both):**
Install the [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)

## Building

```powershell
dotnet build -c Release
```

Generates `SKILL.md` and copies `shared/` references into each skill's `references/` folder.

## Structure

```
skills/
├── shared/          # Shared behavioral guidance (source of truth)
├── excel-mcp/       # MCP Server skill (SKILL.md + references/)
├── excel-cli/       # CLI skill (SKILL.md + references/)
├── templates/       # Scriban templates for SKILL.md generation
├── CLAUDE.md        # Claude Code project instructions
└── .cursorrules     # Cursor-specific rules
```

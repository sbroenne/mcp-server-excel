# Excel MCP Server - Agent Skills

Two skill packages for AI coding assistants:

| Skill | Target | Best For |
|-------|--------|----------|
| **[excel-cli](excel-cli/SKILL.md)** | CLI Tool | Coding agents - token-efficient, `--help` discoverable |
| **[excel-mcp](excel-mcp/SKILL.md)** | MCP Server | Conversational AI - rich tool schemas |

## Installation

```bash
# Via VS Code extension (auto-installs excel-mcp)
# Or via npx:
npx skills add sbroenne/mcp-server-excel --skill excel-cli   # Coding agents
npx skills add sbroenne/mcp-server-excel --skill excel-mcp   # Conversational AI
```

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

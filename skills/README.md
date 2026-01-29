# Excel MCP Server - Agent Skills

Cross-platform AI assistant guidance for Excel MCP Server, following the emerging Agent Skills specification.

## Two Skills for Different Use Cases

| Skill | Target | Best For |
|-------|--------|----------|
| **[excel-cli](excel-cli/SKILL.md)** | CLI Tool | **Coding agents** (Copilot, Cursor, Windsurf) - token-efficient, `--help` discoverable |
| **[excel-mcp](excel-mcp/SKILL.md)** | MCP Server | **Conversational AI** (Claude Desktop, VS Code Chat) - rich tool schemas, exploratory workflows |

### When to use CLI (Recommended for Coding Agents)

Modern coding agents increasingly favor CLI-based workflows over MCP because:
- **Token-efficient**: No large tool schemas loaded into context
- **Discoverable**: Agents can run `excelcli --help` to learn commands
- **Scriptable**: Works in bash pipelines, CI/CD, batch processing
- **Quiet mode**: `-q` flag outputs clean JSON only

```bash
# Coding agent workflow
excelcli -q session open C:\Data\Report.xlsx
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values-json '[["Hello"]]'
excelcli -q session close --session 1 --save
```

### When to use MCP Server

MCP remains relevant for:
- Exploratory automation with iterative reasoning
- Self-healing workflows needing rich introspection  
- Long-running autonomous tasks with continuous context
- Conversational interfaces (Claude Desktop, VS Code Chat)

## What Are Agent Skills?

Agent Skills are reusable instruction sets that extend AI coding assistants with domain-specific knowledge. They enable consistent, reliable behavior when working with specific tools like Excel MCP Server.

## Supported Platforms

| Platform | Install Location | Auto-Install |
|----------|------------------|--------------|
| **GitHub Copilot** | `~/.copilot/skills/excel-mcp/` | Via VS Code extension |
| **Claude Code** | `.claude/skills/excel-mcp/` | Manual or npx |
| **Cursor** | `.cursor/skills/excel-mcp/` | Manual or npx |
| **Windsurf** | `.windsurf/skills/excel-mcp/` | Manual or npx |
| **Gemini CLI** | `.gemini/skills/excel-mcp/` | Manual or npx |
| **Goose** | `.goose/skills/excel-mcp/` | Manual or npx |

## Installation Methods

### Method 1: VS Code Extension (Copilot)

Install the [Excel MCP Server extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) - skills are installed automatically to `~/.copilot/skills/`.

Enable skills in VS Code settings:
```json
{
  "chat.useAgentSkills": true
}
```

### Method 2: npx add-skill (Cross-Platform)

```bash
# Install for all supported agents
npx add-skill sbroenne/mcp-server-excel

# Install for specific agent
npx add-skill sbroenne/mcp-server-excel -a claude-code
npx add-skill sbroenne/mcp-server-excel -a cursor

# Install globally (user-wide)
npx add-skill sbroenne/mcp-server-excel --global

# Install to current project
npx add-skill sbroenne/mcp-server-excel --local
```

### Method 3: GitHub Release Download

1. Download `excel-mcp-skills-vX.X.X.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract to the appropriate directory for your AI assistant:
   - Copilot: `~/.copilot/skills/excel-mcp/`
   - Claude Code: `.claude/skills/excel-mcp/`
   - Cursor: `.cursor/skills/excel-mcp/`
   - Windsurf: `.windsurf/skills/excel-mcp/`

### Method 4: Git Clone (Development)

```bash
# Clone and copy
git clone https://github.com/sbroenne/mcp-server-excel.git
cp -r mcp-server-excel/skills/excel-mcp ~/.copilot/skills/
```

## Directory Structure

```
skills/
├── README.md                    # This file
├── excel-mcp/                   # MCP Server skill package
│   ├── SKILL.md                 # Primary skill definition (MCP tools)
│   ├── VERSION                  # Version tracking
│   └── references/              # Supporting documentation
│       ├── workflows.md         # Production workflow patterns
│       ├── behavioral-rules.md  # Execution guidelines
│       ├── anti-patterns.md     # Common mistakes to avoid
│       ├── excel_powerquery.md  # Power Query specifics
│       ├── excel_datamodel.md   # Data Model/DAX specifics
│       ├── excel_table.md       # Table operations
│       ├── excel_range.md       # Range operations
│       ├── excel_worksheet.md   # Worksheet operations
│       └── claude-desktop.md    # Claude Desktop setup
├── excel-cli/                   # CLI skill package (NEW)
│   └── SKILL.md                 # CLI commands documentation
├── CLAUDE.md                    # Claude Code project instructions
└── .cursorrules                 # Cursor-specific rules
```

## Platform-Specific Files

### For Claude Code Users

Copy `CLAUDE.md` to your project root:
```bash
cp skills/CLAUDE.md /path/to/your/project/CLAUDE.md
```

Or reference from `.claude/skills/`:
```bash
mkdir -p .claude/skills
cp -r skills/excel-mcp .claude/skills/
```

### For Cursor Users

Copy `.cursorrules` to your project root:
```bash
cp skills/.cursorrules /path/to/your/project/.cursorrules
```

## Building the Skills Package

For maintainers building release artifacts:

```powershell
# Build all skill artifacts
./scripts/Build-AgentSkills.ps1

# Output:
#   artifacts/excel-mcp-skills.zip        - Full skill package
#   artifacts/CLAUDE.md                   - Claude Code instructions
#   artifacts/.cursorrules                - Cursor rules
```

## Version Compatibility

| Skills Version | MCP Server Version | Minimum Excel |
|----------------|-------------------|---------------|
| 1.2.0+ | 1.2.0+ | Excel 2016+ |
| 1.1.x | 1.1.x | Excel 2016+ |

## Contributing

See [CONTRIBUTING.md](../docs/CONTRIBUTING.md) for guidelines on improving the skills.

## Related Resources

- [Excel MCP Server Documentation](https://excelmcpserver.dev/)
- [MCP Registry](https://mcp.run/registry)
- [agentskills.io Specification](https://agentskills.io)

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
- **Scriptable**: Works in PowerShell pipelines, CI/CD, batch processing
- **Quiet mode**: `-q` flag outputs clean JSON only

```powershell
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

Download the skill package for your use case from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest):

| Package | Best For |
|---------|----------|
| `excel-cli-skill-vX.X.X.zip` | **Coding agents** (Copilot, Cursor, Windsurf) |
| `excel-mcp-skill-vX.X.X.zip` | **Conversational AI** (Claude Desktop, VS Code Chat) |

Extract to the appropriate directory for your AI assistant:
- Copilot: `~/.copilot/skills/excel-mcp/` or `~/.copilot/skills/excel-cli/`
- Claude Code: `.claude/skills/excel-mcp/` or `.claude/skills/excel-cli/`
- Cursor: `.cursor/skills/excel-mcp/` or `.cursor/skills/excel-cli/`
- Windsurf: `.windsurf/skills/excel-mcp/` or `.windsurf/skills/excel-cli/`

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
├── shared/                      # Shared behavioral guidance (source of truth)
│   ├── behavioral-rules.md      # Core execution rules
│   ├── anti-patterns.md         # Common mistakes to avoid
│   ├── workflows.md             # Production workflow patterns
│   ├── excel_powerquery.md      # Power Query specifics
│   ├── excel_datamodel.md       # Data Model/DAX specifics
│   ├── excel_table.md           # Table operations
│   ├── excel_range.md           # Range operations
│   ├── excel_worksheet.md       # Worksheet operations
│   ├── excel_chart.md           # Chart operations
│   ├── excel_slicer.md          # Slicer operations
│   └── excel_conditionalformat.md # Conditional formatting
├── excel-mcp/                   # MCP Server skill package
│   ├── SKILL.md                 # Primary skill definition (MCP tools)
│   ├── README.md                # MCP skill installation guide
│   ├── VERSION                  # Version tracking
│   └── references/              # MCP-specific + shared (copied during build)
│       ├── claude-desktop.md    # Claude Desktop setup (MCP-specific)
│       └── (shared files)       # Copied from shared/ by build script
├── excel-cli/                   # CLI skill package
│   ├── SKILL.md                 # CLI commands documentation
│   ├── README.md                # CLI skill installation guide
│   └── references/              # CLI-specific + shared (copied during build)
│       ├── README.md            # Explains build process
│       └── (shared files)       # Copied from shared/ by build script
├── CLAUDE.md                    # Claude Code project instructions
└── .cursorrules                 # Cursor-specific rules
```

**Note:** Shared behavioral guidance lives in `shared/` and is copied to each skill's `references/` folder during packaging. For local development, run `./scripts/Build-AgentSkills.ps1 -PopulateReferences` to populate the references folders.

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

## Building the Skills Packages

For maintainers building release artifacts:

```powershell
# Build all skill artifacts (for release)
./scripts/Build-AgentSkills.ps1

# Output:
#   artifacts/skills/excel-mcp-skill-v{version}.zip  - MCP Server skill package
#   artifacts/skills/excel-cli-skill-v{version}.zip  - CLI skill package
#   artifacts/skills/CLAUDE.md                       - Claude Code instructions
#   artifacts/skills/.cursorrules                    - Cursor rules

# For local development: populate references from shared/
./scripts/Build-AgentSkills.ps1 -PopulateReferences
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

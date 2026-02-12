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

### Method 2: npx skills add (Cross-Platform)

The repository contains TWO skills. The CLI will prompt you to select which one(s) to install:

```bash
# Interactive install - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel

# Install specific skill directly
npx skills add sbroenne/mcp-server-excel --skill excel-cli   # Coding agents
npx skills add sbroenne/mcp-server-excel --skill excel-mcp   # Conversational AI

# Install both skills
npx skills add sbroenne/mcp-server-excel --skill '*'

# Install for specific agent
npx skills add sbroenne/mcp-server-excel --skill excel-cli -a cursor
npx skills add sbroenne/mcp-server-excel --skill excel-mcp -a claude-code

# Install globally (user-wide)
npx skills add sbroenne/mcp-server-excel --skill excel-cli --global
```

### Method 3: GitHub Release Download

Download `excel-skills-vX.X.X.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest).

The package contains both skills:
- `skills/excel-cli/` - for coding agents (Copilot, Cursor, Windsurf)
- `skills/excel-mcp/` - for conversational AI (Claude Desktop, VS Code Chat)

Extract the skill(s) you need to your AI assistant's skills directory:
- Copilot: `~/.copilot/skills/excel-cli/` or `~/.copilot/skills/excel-mcp/`
- Claude Code: `.claude/skills/excel-cli/` or `.claude/skills/excel-mcp/`
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
│   ├── powerquery.md        # Power Query specifics
│   ├── datamodel.md         # Data Model/DAX specifics
│   ├── table.md             # Table operations
│   ├── range.md             # Range operations
│   ├── worksheet.md         # Worksheet operations
│   ├── chart.md             # Chart operations
│   ├── slicer.md            # Slicer operations
│   └── conditionalformat.md # Conditional formatting
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

**Note:** Shared behavioral guidance lives in `shared/` and is copied to each skill's `references/` folder during the build. Run `dotnet build -c Release` to generate SKILL.md and copy references.

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

Skill files are generated automatically during Release builds:

```powershell
# Build solution - generates SKILL.md and copies references
dotnet build -c Release

# Output (in skills/ folder):
#   excel-mcp/SKILL.md        - Generated MCP skill documentation
#   excel-mcp/references/     - Copied shared reference files
#   excel-cli/SKILL.md        - Generated CLI skill documentation  
#   excel-cli/references/     - Copied shared reference files
```

For release artifacts (ZIP package), the GitHub Actions workflow handles packaging.

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

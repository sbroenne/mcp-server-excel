# Installation Guide - ExcelMcp

Complete installation instructions for the ExcelMcp MCP Server and CLI tool.

## System Requirements

### Required
- **Windows OS** (Windows 10 or later)
- **Microsoft Excel 2016 or later** (Desktop version - Office 365, Professional Plus, or Standalone)
- **.NET 10 Runtime or SDK** (not required for VS Code Extension or MCPB - they bundle it)

### Optional (for specific features)
- **Microsoft Analysis Services OLE DB Provider (MSOLAP)** - Required for DAX query execution (`evaluate`, `execute-dmv` actions)
  - Easiest: Install [Power BI Desktop](https://powerbi.microsoft.com/desktop) (includes MSOLAP)
  - Alternative: [Microsoft OLE DB Driver for Analysis Services](https://learn.microsoft.com/analysis-services/client-libraries)
- **Node.js** - Only required for `npx` commands (`add-mcp` auto-configuration, agent skills). Install with `winget install OpenJS.NodeJS.LTS` or from [nodejs.org](https://nodejs.org/)

### Recommended
- Windows 11 for best performance
- Office 365 with latest updates
- 8GB RAM minimum

---

## Quick Start (Recommended)

### VS Code Extension (Easiest - One-Click Setup)

**Best for:** GitHub Copilot users, beginners, anyone wanting automatic configuration

1. **Install the Extension**
   - Open VS Code
   - Press `Ctrl+Shift+X` (Extensions)
   - Search for **"ExcelMcp"**
   - Click **Install**

2. **That's It!**
   - Bundles self-contained MCP server and CLI (no .NET runtime or SDK needed)
   - Auto-configures GitHub Copilot
   - Registers agent skills (excel-mcp + excel-cli) via `chatSkills`
   - Shows quick start guide on first launch

**Marketplace Link:** [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)

---

### Claude Desktop (One-Click Install)

**Best for:** Claude Desktop users who want the simplest installation

1. Download `excel-mcp-{version}.mcpb` from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Double-click the `.mcpb` file (or drag-and-drop onto Claude Desktop)
3. Restart Claude Desktop

That's it! The MCPB bundle includes everything needed - no .NET installation required.

---

## Manual Installation  & Configuration (MCP Server)

**Best for:** Other MCP clients (Cursor, Windsurf, Cline, Claude Code, Codex), advanced users

### Step 1: Install .NET 10

**Check if already installed:**
```powershell
dotnet --version
# Should show 10.0.x or higher
```

**If not installed:**

```powershell
winget install Microsoft.DotNet.Runtime.10
```

**Manual Download:** [.NET 10 Downloads](https://dotnet.microsoft.com/download/dotnet/10.0)

### Step 2: Install ExcelMcp MCP Server

```powershell
# Install globally as a .NET tool
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Verify installation
dotnet tool list --global | Select-String "ExcelMcp"
```

### Step 3: Configure Your MCP Client

#### Option A: Auto-Configure All Agents (Recommended)

Use [`add-mcp`](https://github.com/neondatabase/add-mcp) to configure all detected coding agents with a single command:

```powershell
npx add-mcp "mcp-excel" --name excel-mcp
```

This auto-detects and configures **Cursor, VS Code, Claude Code, Claude Desktop, Codex, Zed, Gemini CLI**, and more. Use flags to customize:

```powershell
# Configure specific agents only
npx add-mcp "mcp-excel" --name excel-mcp -a cursor -a claude-code

# Configure globally (user-wide, all projects)
npx add-mcp "mcp-excel" --name excel-mcp -g

# Non-interactive (skip prompts)
npx add-mcp "mcp-excel" --name excel-mcp --all -y
```

> **Requires:** [Node.js](https://nodejs.org/) for `npx`. Install with `winget install OpenJS.NodeJS.LTS` if not already available. No permanent `add-mcp` installation needed — `npx` downloads, runs, and cleans up automatically.

#### Option B: Manual Configuration

**Quick Start:** Ready-to-use config files for all clients are available in [`examples/mcp-configs/`](https://github.com/sbroenne/mcp-server-excel/tree/main/examples/mcp-configs/)

**For GitHub Copilot (VS Code):**

Create `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel-mcp": {
      "command": "mcp-excel"
    }
  }
}
```

**For GitHub Copilot (Visual Studio):**

Create `.mcp.json` in your solution directory or `%USERPROFILE%\.mcp.json`:

```json
{
  "servers": {
    "excel-mcp": {
      "command": "mcp-excel"
    }
  }
}
```

**For Claude Desktop:**

1. Locate config file: `%APPDATA%\Claude\claude_desktop_config.json`
2. If file doesn't exist, create it with the content below
3. If file exists, merge the `excel-mcp` entry into your existing `mcpServers` section

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "mcp-excel",
      "args": [],
      "env": {}
    }
  }
}
```

4. Save and restart Claude Desktop

**For Cursor:**

1. Open Cursor Settings (Ctrl+,)
2. Search for "MCP" in settings
3. Click "Edit in settings.json" or create config at: `%APPDATA%\Cursor\User\globalStorage\mcp\mcp.json`
4. Add this configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "mcp-excel",
      "args": [],
      "env": {}
    }
  }
}
```

5. Save and restart Cursor

**For Cline (VS Code Extension):**

1. Install Cline extension in VS Code
2. Open Cline panel and click the MCP settings gear icon
3. Add this configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "mcp-excel",
      "args": [],
      "env": {}
    }
  }
}
```

4. Save and restart VS Code

**For Windsurf:**

1. Open Windsurf Settings
2. Navigate to MCP Servers configuration  
3. Add this configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "mcp-excel",
      "args": [],
      "env": {}
    }
  }
}
```

4. Save and restart Windsurf

### Step 4: Test the Installation

Restart your MCP client, then ask:
```
Create an empty Excel file called "test.xlsx"
```

If it works, you're all set! 🎉

**💡 Tip:** Want to watch the AI work? Ask:
```
Show me Excel while you work on test.xlsx
```
This opens Excel visibly so you can see every change in real-time - great for debugging and demos!

---

## CLI Installation (No AI Required)

**Best for:** Scripting, RPA, CI/CD pipelines, automation without AI

> **📦 Bundled with MCP Server:** The CLI (`excelcli`) is included in the unified package. Install once, get both tools!

### Install Unified Package

```powershell
# Install unified package (includes MCP Server + CLI)
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Verify CLI is available
excelcli --version
```

### Quick Test

```powershell
# Session-based workflow (keeps Excel open between commands)
excelcli -q session open test.xlsx        # Returns session ID
excelcli -q sheet list --session 1        # List worksheets
excelcli -q session close --session 1 --save
```

> **💡 Tip:** Use `-q` (quiet mode) to suppress banner and get JSON output only - perfect for scripting and automation.

**CLI Documentation:** [CLI Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)

---

## Agent Skills Installation (Cross-Platform)

**Best for:** Adding AI guidance to coding agents (Copilot, Cursor, Windsurf, Claude Code, Gemini, Codex, etc.)

Skills are auto-installed by the VS Code extension. For other platforms:

```powershell
# CLI skill (for coding agents - token-efficient workflows)
npx skills add sbroenne/mcp-server-excel --skill excel-cli

# MCP skill (for conversational AI - rich tool schemas)
npx skills add sbroenne/mcp-server-excel --skill excel-mcp

# Install for specific agents
npx skills add sbroenne/mcp-server-excel --skill excel-cli -a cursor
npx skills add sbroenne/mcp-server-excel --skill excel-mcp -a claude-code

# Install globally (user-wide)
npx skills add sbroenne/mcp-server-excel --skill excel-cli --global
```

**Supports 43+ agents** including claude-code, github-copilot, cursor, windsurf, gemini-cli, codex, goose, cline, continue, replit, and more.

**📚 [Agent Skills Guide →](../skills/README.md)**

---

## Updating ExcelMcp

### Check Installed Version

**MCP Server and CLI:**
```powershell
dotnet tool list --global | Select-String "ExcelMcp"

# Or check CLI version
excelcli --version
```

### Update (MCP Server + CLI)

> **📦 Unified Package:** Updating the MCP Server also updates the CLI - they're bundled together!

**Step 1: Update the tool**
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
```

**Step 2: Verify update**
```powershell
# Check installed version
dotnet tool list --global | Select-String "ExcelMcp"

# Verify both tools work
excelcli --version
mcp-excel --version
```

**Step 3: Restart your MCP client**
- Restart VS Code, Claude Desktop, Cursor, or whichever client you're using
- The new version will be used automatically

### Troubleshooting Updates

#### Update Command Fails

**Error: "Tool not found"**
```powershell
# The tool may need to be reinstalled
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

**Error: "Access denied"**
- Run PowerShell as Administrator
- Or install in user directory (not global):
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer --install-dir ~/.dotnet/tools
```

#### MCP Server Still Running Old Version

**Solution:** Fully restart your MCP client
- Close VS Code completely (including terminal windows)
- Close Claude Desktop completely
- Reopen the application

**Still not working?**
```powershell
# Reinstall the tool
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

### Rollback to Previous Version

If an update causes issues, you can downgrade:

```powershell
# Uninstall current version
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer

# Install specific version
dotnet tool install --global Sbroenne.ExcelMcp.McpServer --version 1.2.3
# Replace 1.2.3 with the version you want
```

### Check What's New

Before updating, check the release notes:
- **GitHub Releases:** https://github.com/sbroenne/mcp-server-excel/releases
- **Changelog:** https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md

---

## Troubleshooting

### Common Issues

#### 1. "dotnet command not found"

**Solution:** Install .NET 10 SDK or Runtime (see Step 1 above)

#### 2. MCP Server Not Responding

**Check if tool is installed:**
```powershell
dotnet tool list --global | Select-String "ExcelMcp"
```

**Reinstall if missing:**
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

#### 3. "Workbook is locked" or "Cannot open file"

**Solution:** Close all Excel windows before running ExcelMcp

ExcelMcp requires exclusive access to workbooks (Excel COM limitation).

## Uninstallation

### Uninstall MCP Server
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
```

### Uninstall CLI
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

---

## Getting Help

- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md)

---

## Next Steps

After installation:

1. **Learn the basics:** Try simple commands like creating worksheets, setting values
2. **Explore features:** See [README](https://github.com/sbroenne/mcp-server-excel#readme) for complete feature list
3. **Read the guides:**
   - [MCP Server Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.McpServer/README.md)
   - [CLI Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)
   - [Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-mcp/SKILL.md) - Cross-platform AI guidance
4. **Join the community:** Star the repo, report issues, contribute improvements

---

## Agent Skills (Optional)

Agent Skills provide domain-specific guidance to AI coding assistants, helping them use Excel MCP Server more effectively.

> **Note:** Agent Skills are for **coding agents** (GitHub Copilot, Claude Code, Cursor). **Claude Desktop** uses MCP Prompts instead (included automatically via the MCP Server).

### Two Skills for Different Use Cases

| Skill | Target | Best For |
|-------|--------|----------|
| **excel-cli** | CLI Tool | **Coding agents** (Copilot, Cursor, Windsurf) - token-efficient, `excelcli --help` discoverable |
| **excel-mcp** | MCP Server | **Conversational AI** (Claude Desktop, VS Code Chat) - rich tool schemas, exploratory workflows |

**VS Code Extension:** Skills are installed automatically to `~/.copilot/skills/`.

**Other Platforms (Claude Code, Cursor, Windsurf, Gemini, Codex, etc.):**

```powershell
# Install CLI skill (recommended for coding agents - Copilot, Cursor, Windsurf, Codex, etc.)
npx skills add sbroenne/mcp-server-excel --skill excel-cli

# Install MCP skill (for conversational AI - Claude Desktop, VS Code Chat)
npx skills add sbroenne/mcp-server-excel --skill excel-mcp

# Interactive install - prompts to select excel-cli, excel-mcp, or both
npx skills add sbroenne/mcp-server-excel

# Install specific skill directly
npx skills add sbroenne/mcp-server-excel --skill excel-cli   # Coding agents
npx skills add sbroenne/mcp-server-excel --skill excel-mcp   # Conversational AI

# Install both skills
npx skills add sbroenne/mcp-server-excel --skill '*'

# Target specific agent (optional - auto-detects if omitted)
npx skills add sbroenne/mcp-server-excel --skill excel-cli -a cursor
npx skills add sbroenne/mcp-server-excel --skill excel-mcp -a claude-code
```

**Manual Installation:**
1. Download `excel-skills-v{version}.zip` from [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. The package contains both skills:
   - `skills/excel-cli/` - for coding agents (Copilot, Cursor, Windsurf)
   - `skills/excel-mcp/` - for conversational AI (Claude Desktop, VS Code Chat)
3. Extract the skill(s) you need to your AI assistant's skills directory:
   - Copilot: `~/.copilot/skills/excel-cli/` or `~/.copilot/skills/excel-mcp/`
   - Claude Code: `.claude/skills/excel-cli/` or `.claude/skills/excel-mcp/`
   - Cursor: `.cursor/skills/excel-cli/` or `.cursor/skills/excel-mcp/`

**See:** [Agent Skills Documentation](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/README.md)

---

**Happy automating! 🚀**

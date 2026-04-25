# Installation Guide - ExcelMcp

Complete installation instructions for the ExcelMcp MCP Server and CLI tool.

## System Requirements

### Required
- **Windows OS** (Windows 10 or later)
- **Microsoft Excel 2016 or later** (Desktop version - Office 365, Professional Plus, or Standalone)

> **.NET runtime is NOT required** for any installation method — all distributions are self-contained.

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

Use this order to avoid setup confusion:

1. **Choose one primary setup path**:
   - **VS Code Extension** (GitHub Copilot users) — auto-configures everything
   - **Claude Desktop MCPB** — one-click MCP installation
   - **GitHub Copilot Plugins** (Copilot CLI users) — marketplace installation
   - **Manual MCP setup** (other MCP clients like Cursor, Windsurf)
2. **Validate MCP setup** (run the quick test prompt in Step 4 of manual setup, or test in your client after extension/MCPB/plugin install)
3. **Optional:** install CLI (`excelcli`) for scripting/RPA (auto-included with Copilot CLI plugins)
4. **Optional:** install agent skills separately for non-extension environments

### GitHub Copilot Plugins (Easiest for Copilot CLI Users)

**Best for:** GitHub Copilot CLI users who want plugin marketplace installation

ExcelMcp is published as two complementary plugins in the GitHub Copilot plugin marketplace:

**Excel MCP** (25 tools with 230 operations for conversational AI):
```powershell
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@sbroenne/mcp-server-excel-plugins
```

**Excel CLI** (Skill + bundled CLI for coding agents):
```powershell
copilot plugin install excel-cli@sbroenne/mcp-server-excel-plugins
pwsh -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
```

Both plugins are maintained in [`sbroenne/mcp-server-excel-plugins`](https://github.com/sbroenne/mcp-server-excel-plugins) and auto-updated after each release.

---

**Best for:** GitHub Copilot users, beginners, anyone wanting automatic configuration

1. **Install the Extension**
   - Open VS Code
   - Press `Ctrl+Shift+X` (Extensions)
   - Search for **"ExcelMcp"**
   - Click **Install**

2. **That's It!**
   - Bundles self-contained MCP server and CLI (no .NET runtime needed)
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

## Manual MCP Setup (All MCP Clients)

**Best for:** Other MCP clients (Cursor, Windsurf, Cline, Claude Code, Codex), advanced users

### Step 1: Download MCP Server

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download **`ExcelMcp-MCP-Server-{version}-windows.zip`**
3. Extract the ZIP to a permanent location (e.g., `C:\Tools\ExcelMcp\`)

```powershell
# Example extraction
Expand-Archive "ExcelMcp-MCP-Server-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp"
```

The ZIP contains `mcp-excel.exe` — a fully self-contained executable (no .NET runtime needed).

### Step 2: Add to PATH (Recommended)

To use `mcp-excel` as a command without specifying the full path:

```powershell
# Add to user PATH (persistent)
$toolsDir = "C:\Tools\ExcelMcp"
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*$toolsDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$userPath;$toolsDir", "User")
    Write-Host "Added $toolsDir to user PATH. Restart your terminal to apply."
}
```

Or manually: **Settings → System → About → Advanced system settings → Environment Variables → User variables → Path → Edit → New** → add `C:\Tools\ExcelMcp`

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

> **Note:** If `mcp-excel` is not on your PATH, use the full path instead: `npx add-mcp "C:\Tools\ExcelMcp\mcp-excel.exe" --name excel-mcp`

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

> If `mcp-excel` is not on PATH, use the full path: `"command": "C:\\Tools\\ExcelMcp\\mcp-excel.exe"`

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

### Step 4: Validate MCP Setup

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

## Optional: CLI Installation (No AI Required)

**Best for:** Scripting, RPA, CI/CD pipelines, automation without AI

The `excelcli.exe` tool is already included when you install the **excel-cli GitHub Copilot plugin** or the VS Code extension. Plain skill-only installs still need the CLI available separately.

### If Not Already Installed Via Plugin or VS Code Extension

Download and extract the standalone CLI:

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download **`ExcelMcp-CLI-{version}-windows.zip`**
3. Extract to a permanent location (e.g., `C:\Tools\ExcelMcp\`)

```powershell
Expand-Archive "ExcelMcp-CLI-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp"
```

### Add CLI to PATH

```powershell
$toolsDir = "C:\Tools\ExcelMcp"
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*$toolsDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$userPath;$toolsDir", "User")
    Write-Host "Added $toolsDir to user PATH. Restart your terminal to apply."
}
```

### Quick Test

```powershell
excelcli --version
excelcli --help

# Test with a session
excelcli -q session open test.xlsx
excelcli -q session list
excelcli -q session close --session <id>
```

---

## GitHub Copilot Plugins (Alternative Installation)

**Best for:** GitHub Copilot users who want packaged Excel automation plugins through supported plugin surfaces

ExcelMcp ships two **GitHub Copilot marketplace plugins**:

- **`excel-mcp`** — Best for conversational Excel workflows through the MCP server
- **`excel-cli`** — Best for token-efficient scripting and coding-agent workflows
- You can install **either plugin alone or both together**

### Copilot CLI install path

This is the documented install flow for the published marketplace repo:

```powershell
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install one or both plugins
copilot plugin install excel-mcp@mcp-server-excel-plugins
copilot plugin install excel-cli@mcp-server-excel-plugins
```

### One-time post-install steps

- **`excel-mcp`** — follow the plugin README if you want to merge the bundled MCP config into your user-level Copilot config.
- **`excel-cli`** — the plugin now ships a self-contained `excelcli.exe` deliverable (no .NET runtime required). Run the bundled helper once to expose it on PATH:

```powershell
pwsh -ExecutionPolicy Bypass -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
excelcli --version
```

### Other supported plugin surfaces

- **VS Code agent plugins** — VS Code supports agent plugins in preview. See the official docs for supported plugin behavior and enablement: [Agent plugins in VS Code](https://code.visualstudio.com/docs/copilot/customization/agent-plugins) and [third-party agents in VS Code](https://code.visualstudio.com/docs/copilot/agents/third-party-agents).
- **Claude plugin system** — Claude supports plugins with its own component model and install/runtime rules. See the official [Claude plugins reference](https://code.claude.com/docs/en/plugins-reference).

> **Important:** The commands above are the **GitHub Copilot CLI** install commands for the published marketplace repo `sbroenne/mcp-server-excel-plugins`. Do not assume the same commands apply to VS Code or Claude. Use the surface-specific docs for those environments.

> **Source-layout note:** This source repo is not itself a Copilot CLI marketplace. The `.github/plugins/` folders are source-owned overlays that the publish workflow copies into the published marketplace repo.

The plugins are republished automatically after every successful ExcelMcp release by a follow-up workflow that uses a stored cross-repo PAT scoped to the published marketplace repo. That publish path is sync-gated (so unchanged plugin-facing releases do not force a republish), keeps downgrade/tag mismatch guards in place, stages the self-contained `excelcli.exe` publish output into the `excel-cli` plugin, and retains a manual maintainer re-sync path for repair/replay scenarios. If a fresh GitHub release is visible but the plugin marketplace is still catching up, wait for the follow-up **Publish Plugins** workflow to finish.

---

## Alternative: NuGet .NET Tool Installation (Secondary)

**For users who prefer package managers or already have .NET installed**

NuGet is a secondary distribution channel. It requires the **.NET 10 Runtime or SDK** to be installed.

```powershell
# Requires .NET 10 Runtime or SDK
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

After installation, configure your MCP client with `"command": "mcp-excel"` (same as standalone exe).

**Update via NuGet:**
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

**Uninstall:**
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

> **Why NuGet is secondary:** The standalone exe distributions require no .NET runtime, making them easier to install for most users. NuGet is available as an alternative for users who prefer package managers or already have .NET installed in their workflow.

---

## Agent Skills Installation (Cross-Platform)

**Best for:** Adding AI guidance to coding agents (Copilot, Cursor, Windsurf, Claude Code, Gemini, Codex, etc.)

Skills are auto-installed by the VS Code extension. Plugins and skills are different things: plugins are packaged surface integrations, while skills are reusable AI guidance. For environments where you want skills directly, use the commands below:

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

### Check Current Version

```powershell
# Check MCP Server version
mcp-excel --version

# Check CLI version
excelcli --version
```

### Update to New Version

**Standalone exe (primary):**

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download the new ZIP(s): `ExcelMcp-MCP-Server-{version}-windows.zip` and/or `ExcelMcp-CLI-{version}-windows.zip`
3. Extract and overwrite the existing files in your installation directory

```powershell
# Example update
Expand-Archive "ExcelMcp-MCP-Server-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp" -Force
```

4. Restart your MCP client (VS Code, Claude Desktop, Cursor, etc.)

**NuGet (secondary):**

```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Check What's New

Before updating, check the release notes:
- **GitHub Releases:** https://github.com/sbroenne/mcp-server-excel/releases
- **Changelog:** https://github.com/sbroenne/mcp-server-excel/blob/main/CHANGELOG.md

---

## Troubleshooting

### Common Issues

#### 1. "mcp-excel is not recognized as an internal or external command"

**Solution:** `mcp-excel.exe` is not on your PATH.

Either:
- Add the directory containing `mcp-excel.exe` to your PATH (see Step 2 above)
- Or use the full path in your MCP client config: `"command": "C:\\Tools\\ExcelMcp\\mcp-excel.exe"`

#### 2. MCP Server Not Responding

**Check if the exe exists:**
```powershell
where.exe mcp-excel
# Or with full path:
Test-Path "C:\Tools\ExcelMcp\mcp-excel.exe"
```

**Verify it runs:**
```powershell
mcp-excel --version
```

#### 3. "Workbook is locked" or "Cannot open file"

**Solution:** Close all Excel windows before running ExcelMcp

ExcelMcp requires exclusive access to workbooks (Excel COM limitation).

#### 4. MCP Server Still Running Old Version

**Solution:** Fully restart your MCP client
- Close VS Code completely (including terminal windows)
- Close Claude Desktop completely
- Reopen the application

## Uninstallation

### Uninstall MCP Server
```powershell
# Standalone exe: simply delete the extracted files
Remove-Item "C:\Tools\ExcelMcp\mcp-excel.exe" -Force

# Remove from PATH if you added it
# Settings → System → About → Advanced system settings → Environment Variables
# Edit PATH and remove the ExcelMcp directory

# NuGet (if installed via dotnet tool):
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
```

### Uninstall CLI
```powershell
# Standalone exe:
Remove-Item "C:\Tools\ExcelMcp\excelcli.exe" -Force

# NuGet (if installed via dotnet tool):
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

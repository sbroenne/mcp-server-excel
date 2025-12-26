# Installation Guide - ExcelMcp

Complete installation instructions for the ExcelMcp MCP Server and CLI tool.

## System Requirements

### Required
- **Windows OS** (Windows 10 or later)
- **Microsoft Excel 2016 or later** (Desktop version - Office 365, Professional Plus, or Standalone)
- **.NET 8.0 Runtime or SDK**

### Recommended
- Windows 11 for best performance
- Office 365 with latest updates
- 8GB RAM minimum

---

## Quick Start (Recommended)

### Option 1: VS Code Extension (Easiest - One-Click Setup)

**Best for:** GitHub Copilot users, beginners, anyone wanting automatic configuration

1. **Install the Extension**
   - Open VS Code
   - Press `Ctrl+Shift+X` (Extensions)
   - Search for **"ExcelMcp"**
   - Click **Install**

2. **That's It!**
   - Extension automatically installs .NET 8 runtime
   - Bundles MCP server (no separate installation needed)
   - Auto-configures GitHub Copilot
   - Shows quick start guide on first launch

**Marketplace Link:** [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)

---

### Option 2: Manual Installation (MCP Server)

**Best for:** Claude Desktop users, other MCP clients, advanced users

#### Step 1: Install .NET 8

**Check if already installed:**
```powershell
dotnet --version
# Should show 8.0.x or higher
```

**If not installed, choose one:**

**Option A: SDK (Recommended for developers)**
```powershell
winget install Microsoft.DotNet.SDK.8
```

**Option B: Runtime Only (Smaller download)**
```powershell
winget install Microsoft.DotNet.Runtime.8
```

**Manual Download:** [.NET 8 Downloads](https://dotnet.microsoft.com/download/dotnet/8.0)

#### Step 2: Install ExcelMcp MCP Server

```powershell
# Install globally as a .NET tool
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Verify installation
dotnet tool list --global | Select-String "ExcelMcp"
```

#### Step 3: Configure Your MCP Client

**For GitHub Copilot (VS Code):**

Create `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"]
    }
  }
}
```

**For GitHub Copilot (Visual Studio):**

Create `.mcp.json` in your solution directory or `%USERPROFILE%\.mcp.json`:

```json
{
  "servers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"]
    }
  }
}
```

**For Claude Desktop:**

1. Locate config file: `%APPDATA%\Claude\claude_desktop_config.json`
2. If file doesn't exist, create it with the content below
3. If file exists, merge the `excel` entry into your existing `mcpServers` section

```json
{
  "mcpServers": {
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"],
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
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"],
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
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"],
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
    "excel": {
      "command": "dotnet",
      "args": ["tool", "run", "mcp-excel"],
      "env": {}
    }
  }
}
```

4. Save and restart Windsurf

**Quick Copy:** Ready-to-use config files for all clients are available in [`examples/mcp-configs/`](https://github.com/sbroenne/mcp-server-excel/tree/main/examples/mcp-configs/)

#### Step 4: Test the Installation

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

### Option 3: CLI Installation (No AI Required)

**Best for:** Scripting, RPA, CI/CD pipelines, automation without AI

#### Install CLI Tool

```powershell
# Install CLI globally
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# Verify installation
excel-mcp --version
```

#### Quick Test

```powershell
# Create a test workbook
excel-mcp file-create --file "test.xlsx"

# List worksheets
excel-mcp sheet-list --file "test.xlsx"
```

**CLI Documentation:** [CLI Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)

---

## Updating ExcelMcp

### Update MCP Server
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
```

### Update CLI
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Update VS Code Extension
- VS Code automatically notifies you of updates
- Or manually: Extensions → ExcelMcp → Update

---

## Troubleshooting

### Common Issues

#### 1. "dotnet command not found"

**Solution:** Install .NET 8 SDK or Runtime (see Step 1 above)

Verify PATH includes .NET:
```powershell
$env:PATH -split ';' | Select-String "dotnet"
```

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

#### 4. GitHub Copilot Not Finding Server

**Check configuration file exists:**
```powershell
# VS Code
Test-Path ".vscode/mcp.json"

# Visual Studio
Test-Path ".mcp.json"
```

**Restart VS Code/Visual Studio** after creating configuration.

#### 5. Permission Errors on CI/CD

**Solution:** Run with appropriate permissions

```powershell
# Azure DevOps / GitHub Actions
# Ensure runner has Excel installed and user has Excel permissions
```

---

## Advanced Installation Scenarios

### Corporate Environments

**Using internal NuGet feed:**
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer --add-source https://your-feed.com/v3/index.json
```

**Offline installation:**
```powershell
# Download .nupkg file
dotnet tool install --global --add-source ./nupkg Sbroenne.ExcelMcp.McpServer
```

## Uninstallation

### Uninstall MCP Server
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
```

### Uninstall CLI
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

### Remove VS Code Extension
- Extensions → ExcelMcp → Uninstall

### Clean Up Configuration Files
```powershell
# Remove MCP configuration
Remove-Item ".vscode/mcp.json" -ErrorAction SilentlyContinue
Remove-Item ".mcp.json" -ErrorAction SilentlyContinue
Remove-Item "$env:USERPROFILE\.mcp.json" -ErrorAction SilentlyContinue
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
4. **Join the community:** Star the repo, report issues, contribute improvements

**Happy automating! 🚀**

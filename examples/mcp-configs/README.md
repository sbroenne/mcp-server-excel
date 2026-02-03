# MCP Configuration Examples

This directory contains ready-to-use MCP configuration files for various AI coding assistants.

## Quick Setup Guide

### 1. Install ExcelMcp MCP Server

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
```

### 2. Choose Your Client and Copy the Config

Select the configuration file for your AI assistant and follow the instructions below.

---

## Claude Desktop

**Config File:** `claude-desktop-config.json`

**Location:** `%APPDATA%\Claude\claude_desktop_config.json` (Windows)

**Setup Steps:**

1. Open File Explorer and navigate to: `%APPDATA%\Claude\`
2. If `claude_desktop_config.json` doesn't exist, create it
3. Copy the contents of `claude-desktop-config.json` from this folder
4. If you already have a config file, merge the `excel` server entry into your existing `mcpServers` section
5. Restart Claude Desktop

**Test it:**
```
Create an Excel file called "test.xlsx"
```

---

## Cursor

**Config File:** `cursor-mcp-config.json`

**Location:** 
- Windows: `%APPDATA%\Cursor\User\globalStorage\mcp\mcp.json`
- Or: Project-specific `.cursor/mcp.json` in your workspace

**Setup Steps:**

1. Open Cursor Settings (Ctrl+,)
2. Search for "MCP" in settings
3. Click "Edit in settings.json" or manually create the config file at the location above
4. Copy the contents of `cursor-mcp-config.json` from this folder
5. If you already have a config file, merge the `excel` server entry
6. Restart Cursor

**Test it:**
```
Create an Excel file called "test.xlsx"
```

---

## Cline (VS Code Extension)

**Config File:** `cline-mcp-config.json`

**Location:** 
- VS Code User Settings: Click the MCP settings icon in Cline extension
- Or manually: `%APPDATA%\Code\User\globalStorage\saoudrizwan.claude-dev\settings\cline_mcp_settings.json` (Windows)

**Setup Steps:**

1. Install Cline extension in VS Code
2. Open Cline panel
3. Click the MCP settings gear icon
4. Add the server configuration from `cline-mcp-config.json`
5. Restart VS Code

**Test it:**
```
Create an Excel file called "test.xlsx"
```

---

## Windsurf

**Config File:** `windsurf-mcp-config.json`

**Location:** 
- Windows: `%APPDATA%\Windsurf\User\mcp_settings.json`
- Or check Windsurf's MCP settings panel

**Setup Steps:**

1. Open Windsurf Settings
2. Navigate to MCP Servers configuration
3. Add the server configuration from `windsurf-mcp-config.json`
4. Restart Windsurf

**Test it:**
```
Create an Excel file called "test.xlsx"
```

---

## VS Code (GitHub Copilot)

**Config File:** `vscode-mcp-config.json`

**Location:** `.vscode/mcp.json` in your workspace

**Setup Steps:**

**Option A: Use VS Code Extension (Recommended)**
1. Install the [Excel MCP VS Code Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
2. Configuration is automatic!

**Option B: Manual Configuration**
1. Create `.vscode/mcp.json` in your project
2. Copy contents from `vscode-mcp-config.json`
3. Reload VS Code window

**Test it:**
```
Create an Excel file called "test.xlsx"
```

---

## Troubleshooting

### Server Not Responding

1. **Verify installation:**
 ```powershell
 dotnet tool list --global | Select-String "ExcelMcp"
 ```

2. **Check .NET is installed:**
 ```powershell
 dotnet --version
 # Should show 10.0.x or higher
 ```

3. **Reinstall if needed:**
 ```powershell
 dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
 dotnet tool install --global Sbroenne.ExcelMcp.McpServer
 ```

### Excel Not Found

- Ensure Microsoft Excel Desktop (2016+) is installed
- ExcelMcp requires Windows OS with Excel installed

### Permission Issues

- Close all Excel windows before running ExcelMcp
- Ensure your user account has Excel access

### Still Having Issues?

- Check the [main installation guide](../../docs/INSTALLATION.md)
- Report issues on [GitHub](https://github.com/sbroenne/mcp-server-excel/issues)

---

## Configuration Options

### Multiple Workspaces

If you work with multiple workspaces, you can:
- Use project-specific config files (recommended)
- Or use global user-level configuration

---

## Learn More

- **[Main README](../../README.md)** - Feature overview and examples
- **[Installation Guide](../../docs/INSTALLATION.md)** - Comprehensive setup instructions
- **[MCP Server README](../../src/ExcelMcp.McpServer/README.md)** - Tool documentation
- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code and issues

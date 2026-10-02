# MCP Configuration Examples

This directory contains ready-to-use MCP configuration files for various AI coding assistants.

## Quick Setup Guide

### 1. Install Node.js for Direct npx Setup

```powershell
winget install OpenJS.NodeJS.LTS
npx -y @sbroenne/mcp-server-excel@latest --version
```

The supplied configurations use `npx -y @sbroenne/mcp-server-excel@latest`.
Windows, desktop Excel 2016+, and an interactive desktop are required; .NET is
not. Restart your client after installing Node.js so it sees the updated PATH.
Network access is needed for downloads and update checks. `@latest` uses normal
npm caching and does not upgrade an already running server.

For standalone ZIP or NuGet setup, see the
[installation guide](../../docs/INSTALLATION-MCP-SERVER.md) and replace the
example's command with `mcp-excel`, removing its npx arguments.

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
4. If you already have a config file, merge the `excel-mcp` server entry into your existing `mcpServers` section
5. Restart Claude Desktop

**Test it:**
```
Create an Excel file called "test.xlsx"
```

---

## Cursor

**Config File:** `cursor-mcp-config.json`

**Location:** 
- Windows: `%USERPROFILE%\.cursor\mcp.json`
- Or: Project-specific `.cursor/mcp.json` in your workspace

**Setup Steps:**

1. Open Cursor Settings (Ctrl+,)
2. Search for "MCP" in settings
3. Click "Edit in settings.json" or manually create the config file at the location above
4. Copy the contents of `cursor-mcp-config.json` from this folder
5. If you already have a config file, merge the `excel-mcp` server entry
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
- Use **Open MCP config file** in the client's MCP settings; do not create a
  guessed file under `%APPDATA%\Windsurf`.
- In older Windsurf versions, the file is
  `%USERPROFILE%\.codeium\windsurf\mcp_config.json`.

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
   npx -y @sbroenne/mcp-server-excel@latest --version
   ```

2. **Check Node.js/npm is installed:**
   ```powershell
   node --version
   npm --version
   ```

3. **If npx is missing, install Node.js LTS and restart the client:**
   ```powershell
   winget install OpenJS.NodeJS.LTS
   ```

### Excel Not Found

- Ensure Microsoft Excel Desktop (2016+) is installed
- ExcelMcp requires Windows OS with Excel installed

### Permission Issues

- Close all Excel windows before running ExcelMcp
- Ensure your user account has Excel access

### Still Having Issues?

- Check the [MCP Server installation guide](../../docs/INSTALLATION-MCP-SERVER.md)
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
- **[MCP Server Installation Guide](../../docs/INSTALLATION-MCP-SERVER.md)** - Comprehensive setup instructions
- **[MCP Server README](../../src/ExcelMcp.McpServer/README.md)** - Tool documentation
- **[GitHub Repository](https://github.com/sbroenne/mcp-server-excel)** - Source code and issues

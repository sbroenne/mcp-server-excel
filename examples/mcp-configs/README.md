# MCP Configuration Examples

This directory contains ready-to-use MCP configuration files for various AI coding assistants.

**Apple Silicon macOS support is experimental beta**, with only the
[enabled action subset](../../docs/MACOS-ACTION-INVENTORY.md).
See [unsupported features](../../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta).
Windows retains the complete backend. Interactive desktop Excel is required;
Intel Macs, Linux, and headless hosts are unsupported.

## Quick Setup Guide

### 1. Install ExcelMcp MCP Server

These configuration files use the bare `mcp-excel` command, so install the
matching standalone runtime on PATH: download
`ExcelMcp-MCP-Server-{version}-windows.zip` or
`ExcelMcp-MCP-Server-{version}-macos-arm64.zip` from
[Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest),
extract `mcp-excel.exe` or `mcp-excel` to a permanent directory, and add that
directory to PATH. Alternatively, replace the command with its absolute path.

For a direct npm configuration (`npx`, Node.js 18+) or the secondary .NET-tool
channel, follow the [installation guide](../../docs/INSTALLATION-MCP-SERVER.md).
The standalone runtimes do not require .NET or Node.js.

### 2. Choose Your Client and Copy the Config

Select the configuration file for your AI assistant and follow the instructions below.

---

## Claude Desktop

**Config File:** `claude-desktop-config.json`

**Location:** `%APPDATA%\Claude\claude_desktop_config.json` (Windows), or
`~/Library/Application Support/Claude/claude_desktop_config.json` (Mac)

**Setup Steps:**

1. Open the configuration directory for your platform
2. If `claude_desktop_config.json` doesn't exist, create it
3. Copy the contents of `claude-desktop-config.json` from this folder
4. If you already have a config file, merge the `excel-mcp` server entry into your existing `mcpServers` section
5. Restart Claude Desktop

**Test it:**
```
Create an Excel file at this supplied absolute path ending in "test.xlsx"
```

---

## Cursor

**Config File:** `cursor-mcp-config.json`

**Location:** 
- Windows: `%APPDATA%\Cursor\User\globalStorage\mcp\mcp.json`
- Or: Project-specific `.cursor/mcp.json` in your workspace

**Setup Steps:**

1. Open Cursor Settings (Ctrl+, on Windows; Cmd+, on Mac)
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

1. **Verify the executable configured by these examples is on PATH:**
   ```powershell
   mcp-excel --version
   ```

2. **Verify the client's command/path:** GUI clients may have a different PATH
   from your shell; use an absolute executable path if necessary.

3. **Update the matching installation:** replace the standalone runtime in its
   permanent directory. Only .NET-tool installations need the .NET runtime and
   `dotnet tool` diagnostics.

### Excel Not Found

- Ensure desktop Excel 2016+ on Windows x64/ARM64, or Excel for Mac 16.112+ on Apple Silicon
- Verify Excel starts normally in an interactive desktop session

### Permission Issues

- Ensure your user account has Excel and workbook access
- On Mac, grant Excel Automation permission manually when requested
- Reuse the matching workbook session; do not close unrelated workbooks
- For Mac `RecoveryRequired`, reconcile the exact file and pending dialogs
  manually before restarting the client; do not kill shared Excel

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

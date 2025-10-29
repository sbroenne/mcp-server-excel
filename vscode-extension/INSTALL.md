# VS Code Extension - Installation Guide

## Installation Options

### Option 1: VS Code Marketplace (Recommended - Easiest)

**The easiest way to install:**

1. **Open VS Code**
2. **Go to Extensions** panel (`Ctrl+Shift+X` or click Extensions icon in left sidebar)
3. **Search** for "ExcelMcp"
4. **Click Install** on the ExcelMcp extension
5. **Done!** Extension activates automatically

**What happens automatically:**
- ✅ .NET 8 runtime installed (via .NET Install Tool extension)
- ✅ ExcelMcp MCP server tool installed
- ✅ MCP server registered for AI assistants
- ✅ Welcome message shows you're ready

**Marketplace Link:** [ExcelMcp on VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp)

---

### Option 2: Manual VSIX Install

**If you prefer to install from a file:**

1. **Download** `excelmcp-1.0.0.vsix` from the [Releases page](https://github.com/sbroenne/mcp-server-excel/releases)

2. **Install in VS Code:**
   - Open VS Code
   - Press `Ctrl+Shift+P` (Windows/Linux) or `Cmd+Shift+P` (Mac)
   - Type "Install from VSIX"
   - Select the downloaded `excelmcp-1.0.0.vsix` file

3. **Extension activates** and automatically:
   - Installs .NET 8 runtime (if not already installed)
   - Installs ExcelMcp MCP server tool
   - Registers MCP server

---

### Option 3: Open VSX Registry

**For VS Codium or other Open VSX-compatible editors:**

1. Open your editor
2. Go to Extensions
3. Search for "ExcelMcp"
4. Click Install

**Open VSX Link:** [ExcelMcp on Open VSX](https://open-vsx.org/extension/sbroenne/excelmcp)

---

## Requirements

**Must be installed on your system:**
- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed on your system

**Automatically installed by extension:**
- **.NET 8 Runtime** - Extension handles this via .NET Install Tool
- **ExcelMcp MCP Server** - Extension installs this as a global tool

---

## Verifying Installation

**After installation:**

1. **Look for welcome message** - Extension shows a notification on first activation
2. **Ask GitHub Copilot**: "List all available Excel MCP tools"
3. **Expected result**: You should see 10 Excel tools available:
   - excel_powerquery
   - excel_datamodel
   - table
   - excel_range
   - excel_vba
   - excel_connection
   - excel_worksheet
   - excel_parameter
   - excel_file
   - excel_version

---

## What the Extension Does

The ExcelMcp extension automatically registers the ExcelMcp MCP server with VS Code, making Excel automation available to AI assistants like GitHub Copilot.

- ✅ **Zero configuration needed** - Extension handles everything automatically
- ✅ **Automatic .NET installation** - Uses .NET Install Tool extension
- ✅ **Automatic tool installation** - Installs MCP server on activation
- ✅ **Works everywhere** - All VS Code workspaces, no per-workspace config
- ✅ **10 Excel tools** - Power Query, DAX, VBA, Tables, Ranges, and more

---

## Using the Extension

Once installed, you can ask GitHub Copilot to help with Excel tasks:

```
"List all Power Query queries in workbook.xlsx"
"Export all DAX measures to .dax files"  
"Create a new Excel table from range A1:D100"
"Refactor this Power Query M code for better performance"
```

## Available Tools

The MCP server provides **10 specialized tools**:

1. **excel_powerquery** - Power Query M code (11 actions)
2. **excel_datamodel** - DAX measures & relationships (20 actions)
3. **table** - Excel Tables/ListObjects (22 actions)
4. **excel_range** - Range operations (30+ actions)
5. **excel_vba** - VBA macros (7 actions)
6. **excel_connection** - Data connections (11 actions)
7. **excel_worksheet** - Worksheet lifecycle (5 actions)
8. **excel_parameter** - Named ranges (6 actions)
9. **excel_file** - File creation (1 action)
10. **excel_version** - Update checking (1 action)

## Troubleshooting

### Extension not working?

1. **Check .NET 10 is installed**:
   ```powershell
   dotnet --version
   # Should show 10.x.x
   ```

2. **Verify dnx command works**:
   ```powershell
   dnx --help
   ```

3. **Check Excel is installed**:
   - Open Excel manually to verify

4. **Restart VS Code** after installing prerequisites

### Still having issues?

- Check the [Issues page](https://github.com/sbroenne/mcp-server-excel/issues)
- Review the [main documentation](https://github.com/sbroenne/mcp-server-excel)

## Alternative: Manual Configuration

If you prefer not to use the extension, you can manually configure the MCP server by creating `.vscode/mcp.json`:

```json
{
  "servers": {
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer", "--yes"]
    }
  }
}
```

## Development

To build the extension from source:

```bash
cd vscode-extension
npm install
npm run compile
npm run package
```

This creates `excelmcp-1.0.0.vsix` which can be installed in VS Code.

## Publishing

The extension is **automatically published** to the VS Code Marketplace via GitHub Actions when a version tag is pushed (e.g., `vscode-v1.0.0`). 

The release workflow automatically:
- ✅ Updates the version number in `package.json`
- ✅ Updates `CHANGELOG.md` with the release date
- ✅ Publishes to VS Code Marketplace
- ✅ Creates a GitHub release with the VSIX file

**Manual Publishing** (if needed):
See [MARKETPLACE-PUBLISHING.md](MARKETPLACE-PUBLISHING.md) for manual publishing instructions and setup details.

## License

MIT License - see [LICENSE](../LICENSE)

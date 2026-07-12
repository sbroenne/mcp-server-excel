# Installing the CLI - ExcelMcp

Installation instructions for the ExcelMcp **CLI** (`excelcli`) — the entry point for scripting, RPA, CI/CD pipelines, and coding agents that prefer a token-efficient single-tool interface. Looking for the MCP Server instead? See the [MCP Server Installation Guide](https://excelmcpserver.dev/installation-mcp-server/).

## System Requirements

### Required
- **Windows OS** (Windows 10 or later)
- **Microsoft Excel 2016 or later** (Desktop version - Office 365, Professional Plus, or Standalone)

> **.NET runtime is NOT required** for the standalone exe — it's fully self-contained.

### Optional (for specific features)
- **Microsoft Analysis Services OLE DB Provider (MSOLAP)** - Required for DAX query execution (`evaluate`, `execute-dmv` actions)
  - Easiest: Install [Power BI Desktop](https://powerbi.microsoft.com/desktop) (includes MSOLAP)
  - Alternative: [Microsoft OLE DB Driver for Analysis Services](https://learn.microsoft.com/analysis-services/client-libraries)

---

## Quick Start (Recommended)

The **excel-cli GitHub Copilot plugin** bootstraps `excelcli.exe` automatically on first use (downloads and caches the latest release — no separate install needed for plugin-driven flows). The **VS Code extension** does *not* include the CLI (it only bundles the MCP server); install the CLI separately if you need it for scripting. Otherwise:

1. Download and extract the standalone CLI (below)
2. Add it to your PATH
3. Run the quick test to validate

### Standalone Executable (Primary)

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

Or manually: **Settings → System → About → Advanced system settings → Environment Variables → User variables → Path → Edit → New** → add `C:\Tools\ExcelMcp`

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

## GitHub Copilot Plugin

**Best for:** GitHub Copilot CLI users who want token-efficient scripting/skill guidance through the plugin marketplace

```powershell
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install the CLI plugin
copilot plugin install excel-cli@mcp-server-excel-plugins
```

**After Installation:** Install `excelcli` separately if you need it on PATH:

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
excelcli --version
```

> **Note:** The Copilot CLI install command above is specific to the GitHub Copilot plugin marketplace. VS Code and Claude have their own plugin systems with separate installation flows.

Plugins are published automatically after each ExcelMcp release, though you may need to wait a few moments for the update to appear in the marketplace.

---

## Alternative: NuGet .NET Tool Installation (Secondary)

**For users who prefer package managers or already have .NET installed**

NuGet is a secondary distribution channel. It requires the **.NET 10 Runtime or SDK** to be installed.

```powershell
# Requires .NET 10 Runtime or SDK
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

**Update via NuGet:**
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

**Uninstall:**
```powershell
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
```

> **Why NuGet is secondary:** The standalone exe distribution requires no .NET runtime, making it easier to install for most users. NuGet is available as an alternative for users who prefer package managers or already have .NET installed in their workflow.

---

## Updating the CLI

### Check Current Version

```powershell
excelcli --version
```

### Update to New Version

**Standalone exe (primary):**

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download the new ZIP: `ExcelMcp-CLI-{version}-windows.zip`
3. Extract and overwrite the existing files in your installation directory

```powershell
Expand-Archive "ExcelMcp-CLI-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp" -Force
```

**NuGet (secondary):**

```powershell
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Check What's New

Before updating, check the [changelog](https://excelmcpserver.dev/changelog/) or [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases).

---

## Troubleshooting

### Command Not Found After Installation

```powershell
# Check excelcli.exe location
where.exe excelcli

# If not found, ensure the directory containing excelcli.exe is in your PATH
# The default location after extraction might be: C:\Tools\ExcelMcp\
```

### Excel Not Found

```powershell
# Error: "Microsoft Excel is not installed"
# Solution: Install Microsoft Excel (any version 2016+)
```

### VBA Access Denied

VBA commands require **"Trust access to the VBA project object model"** to be enabled manually in Excel:

1. Open Excel
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice

This is a security setting that must be enabled manually. ExcelMcp does not provide a `setup-vba-trust` or `check-vba-trust` command and never modifies Trust Center settings automatically.

Current VBA support is procedural and module-focused:
- `vba list` and `vba view` inspect existing VBA components and procedures
- `vba import` creates a new standard module from inline code or `--vba-code-file`
- `vba update`, `vba delete`, and `vba run` work against existing component/procedure names

For macro-enabled workbooks, use the `.xlsm` extension:

```powershell
excelcli session create macros.xlsm
# Returns session ID (e.g., 1)
excelcli vba import --session 1 --module-name MyModule --vba-code-file code.vba
excelcli session close --session 1 --save
```

### "Workbook is locked" or "Cannot open file"

**Solution:** Close all Excel windows before running ExcelMcp. ExcelMcp requires exclusive access to workbooks (Excel COM limitation).

### Permission Issues

```powershell
# Run PowerShell/CMD as Administrator if you encounter permission errors
# excelcli.exe is a standalone exe - no installation needed
```

---

## Uninstallation

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
- **Contributing:** [Contributing Guide](https://excelmcpserver.dev/contributing/)

---

## Next Steps

After installation:

1. **Learn the basics:** Try `excelcli --help` and open a session against a test workbook
2. **Explore commands:** See the [Feature Reference](https://excelmcpserver.dev/features/) for all 18 command categories
3. **Read the guides:**
   - [MCP Server Installation Guide](https://excelmcpserver.dev/installation-mcp-server/) - for AI assistants like Claude Desktop and Copilot Chat
   - [Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-cli/SKILL.md) - token-efficient AI guidance for coding agents
4. **Join the community:** Star the repo, report issues, contribute improvements

**Happy automating! 🚀**

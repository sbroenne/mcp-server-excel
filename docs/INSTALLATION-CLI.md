# Installing the CLI - ExcelMcp

Installation instructions for the ExcelMcp **CLI** (`excelcli`) — the entry point for scripting, RPA, CI/CD pipelines, and coding agents that prefer a token-efficient single-tool interface. Looking for the MCP Server instead? See the [MCP Server Installation Guide](INSTALLATION-MCP-SERVER.md).

## System Requirements

### Required
- **Windows:** Windows 10 or later with Microsoft Excel 2016 or later
- **macOS:** Apple Silicon Mac with Microsoft Excel for Mac 16.112 or later

Windows provides the complete operation set. The first macOS release is
capability-gated; see [macOS support](../specs/MACOS-SUPPORT.md).

> **.NET runtime is NOT required** for npm or the standalone exe — both use the fully self-contained runtime.

The npm option also requires **Node.js 18 or later**. Install the current LTS
from [nodejs.org](https://nodejs.org/) (or with
`winget install OpenJS.NodeJS.LTS` on Windows). Windows x64/Arm64 (via x64
emulation) and Apple Silicon macOS are supported.

### Optional (for specific features)
- **Microsoft Analysis Services OLE DB Provider (MSOLAP)** - Required for DAX query execution (`evaluate`, `execute-dmv` actions)
  - Easiest: Install [Power BI Desktop](https://www.microsoft.com/en-us/power-platform/products/power-bi/desktop) (includes MSOLAP)
  - Alternative: [Microsoft OLE DB Driver for Analysis Services](https://learn.microsoft.com/analysis-services/client-libraries)

---

## Quick Start (Recommended)

The **excel-cli GitHub Copilot plugin** guides agents to the cross-platform
`@sbroenne/excelcli` npm launcher, which installs the matching optional runtime.
The **VS Code extension**
does *not* include the CLI (it only bundles the MCP server); install the CLI
separately if you need it for scripting outside the plugin. For a direct installation:

Use npm (below) or download the standalone executable if you prefer not to
install Node.js.

### npm (Primary)

```powershell
npx -y @sbroenne/excelcli --version
npx -y @sbroenne/excelcli --help
```

All CLI arguments follow the package name, for example:

```powershell
npx -y @sbroenne/excelcli -q session open "C:\Data\Test.xlsx"
npx -y @sbroenne/excelcli -q session list
npx -y @sbroenne/excelcli -q session close --session <id>
```

For repeated use, install the command on your PATH:

```powershell
npm install --global @sbroenne/excelcli
excelcli --version
```

The launcher installs `@sbroenne/excelcli-win32-x64` on Windows or
`@sbroenne/excelcli-darwin-arm64` on Apple Silicon as an optional dependency.
Do not use `--omit=optional`. It forwards arguments, standard input/output, and
exit codes to the same native `excelcli` executable; session management and
Excel behavior are unchanged.

Avoid installing multiple distributions of `excelcli` on the same PATH. Use
`Get-Command excelcli` in PowerShell to check which installation your shell
will run.

### Standalone Executable (Also Primary)

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download the archive for your platform:
   - Windows: **`ExcelMcp-CLI-{version}-windows.zip`**
   - Apple Silicon macOS: **`ExcelMcp-CLI-{version}-macos-arm64.zip`**
3. Extract to a permanent location.

```powershell
Expand-Archive "ExcelMcp-CLI-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp"
```

```bash
mkdir -p "$HOME/.local/bin"
unzip ExcelMcp-CLI-1.x.x-macos-arm64.zip -d "$HOME/.local/bin"
chmod +x "$HOME/.local/bin/excelcli"
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

# Test with an existing workbook
excelcli -q session open "C:\Data\Test.xlsx"
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

**After installation:** The plugin downloads, caches, and runs `excelcli`
automatically. If you also need `excelcli` directly on your PATH, use the
global npm installation or standalone executable above, or install the secondary NuGet tool when .NET 10 is
available:

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
excelcli --version
```

> **Note:** The Copilot CLI install command above is specific to the GitHub Copilot plugin marketplace. VS Code and Claude have their own plugin systems with separate installation flows.

Plugins are published automatically after each ExcelMcp release, though you may need to wait a few moments for the update to appear in the marketplace.

---

## Alternative: NuGet .NET Tool Installation (Secondary)

**For users who already have .NET installed or prefer .NET tools**

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

> **Why NuGet is secondary:** npm and standalone exe distributions require no separate .NET runtime. NuGet remains available for users who prefer .NET tools.

---

## Updating the CLI

### Check Current Version

```powershell
excelcli --version
```

### Update to New Version

**npm:**

```powershell
# One-off invocation using the latest release:
npx -y @sbroenne/excelcli@latest --version
# Update a global installation:
npm install --global @sbroenne/excelcli@latest
# Uninstall a global installation:
npm uninstall --global @sbroenne/excelcli
```

**Standalone exe (primary):**

1. Go to the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Download the new ZIP for your platform:
   - Windows: `ExcelMcp-CLI-{version}-windows.zip`
   - Apple Silicon macOS: `ExcelMcp-CLI-{version}-macos-arm64.zip`
3. Extract and overwrite the existing files in your installation directory

```powershell
Expand-Archive "ExcelMcp-CLI-1.x.x-windows.zip" -DestinationPath "C:\Tools\ExcelMcp" -Force
```

```bash
unzip -o ExcelMcp-CLI-1.x.x-macos-arm64.zip -d "$HOME/.local/bin/excelmcp"
chmod +x "$HOME/.local/bin/excelmcp/excelcli"
```

**NuGet (secondary):**

```powershell
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

### Check What's New

Before updating, check the [changelog](../CHANGELOG.md) or [GitHub Releases](https://github.com/sbroenne/mcp-server-excel/releases).

---

## Troubleshooting

### npm Runtime Package Missing

If the launcher cannot find `@sbroenne/excelcli-win32-x64` on Windows,
`@sbroenne/excelcli-darwin-arm64` on Apple Silicon, reinstall with optional
dependencies enabled:

```powershell
npm install --global @sbroenne/excelcli --include=optional
```

The npm launcher reports an error on Linux and unsupported architectures rather
than attempting to start Excel.

### Command Not Found After Installation

```powershell
# Check which executable is on PATH
Get-Command excelcli

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

For complete VBA command usage and a macro-enabled workbook example, see
[Automation & Advanced Features](features/AUTOMATION-ADVANCED.md).

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

- **Troubleshooting:** [Troubleshooting](https://excelmcpserver.dev/troubleshooting/) · [FAQ](https://excelmcpserver.dev/faq/)
- **Documentation:** [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- **Issues:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)
- **Contributing:** [Contributing Guide](CONTRIBUTING.md)

---

## Next Steps

After installation:

1. **Learn the basics:** Try `excelcli --help` and open a session against a test workbook
2. **Explore commands:** See the [Feature Reference](../FEATURES.md) for all 31 feature command categories
3. **Read the guides:**
   - [MCP Server Installation Guide](INSTALLATION-MCP-SERVER.md) - for AI assistants like Claude Desktop and Copilot Chat
   - [Agent Skills](../skills/excel-cli/SKILL.md) - token-efficient AI guidance for coding agents
4. **Join the community:** Star the repo, report issues, contribute improvements

**Happy automating! 🚀**

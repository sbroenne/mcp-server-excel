# Installing the CLI - ExcelMcp

Installation instructions for the ExcelMcp **CLI** (`excelcli`) — the entry point for scripting, RPA, CI/CD pipelines, and coding agents that prefer a token-efficient single-tool interface. Looking for the MCP Server instead? See the [MCP Server Installation Guide](INSTALLATION-MCP-SERVER.md).

## System Requirements

### Required
- **Windows:** Windows 10 or later with Microsoft Excel 2016 or later
- **macOS:** Apple Silicon Mac with Microsoft Excel for Mac 16.112 or later

Windows provides the complete operation set. macOS support is **experimental
beta**, not full Windows parity. Power Query, VBA, Data Model/DAX/OLAP, Tables,
PivotTables, charts, slicers, connections, QueryTables, XML Maps, screenshots,
advanced visual formatting, and Python result reads are unavailable. See
[macOS beta limitations](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta)
before installing; test on copies of important workbooks.

> **.NET runtime is NOT required** for npm or the standalone exe — both use the fully self-contained runtime.

The npm option also requires **Node.js 18 or later**. Install the current LTS
from [nodejs.org](https://nodejs.org/) or use
`winget install OpenJS.NodeJS.LTS` on Windows. npm supports Windows x64/Arm64
and Apple Silicon macOS. ARM64 Node.js on Windows selects a native ARM64 runtime.

### Optional (Windows-only features)
- **Microsoft Analysis Services OLE DB Provider (MSOLAP)** - Required for DAX query execution (`evaluate`, `execute-dmv` actions)
  - Easiest: Install [Power BI Desktop](https://www.microsoft.com/en-us/power-platform/products/power-bi/desktop) (includes MSOLAP)
  - Alternative: [Microsoft OLE DB Driver for Analysis Services](https://learn.microsoft.com/analysis-services/client-libraries)

---

## Quick Start (Recommended)

The **excel-cli GitHub Copilot plugin** guides agents to the cross-platform
`@sbroenne/excelcli` npm launcher, which installs the matching optional runtime.
The **VS Code extension**
does *not* include the CLI (it only bundles the MCP server); install the CLI
separately if you need it for scripting outside the plugin. Intel macOS is
unsupported and fails closed. For a direct installation:

Use npm (below) or download the standalone executable if you prefer not to
install Node.js.

### npm (Primary)

```powershell
npx -y @sbroenne/excelcli@latest --version
npx -y @sbroenne/excelcli@latest --help
```

All CLI arguments follow the package name, for example:

```powershell
npx -y @sbroenne/excelcli@latest -q session open "C:\Data\Test.xlsx"
npx -y @sbroenne/excelcli@latest -q session list
npx -y @sbroenne/excelcli@latest -q session close --session <id>
```

The open example uses a Windows path. On Mac, pass your actual absolute path,
for example `npx -y @sbroenne/excelcli@latest -q session open "/absolute/path/Test.xlsx"`.
Capture the returned session ID before follow-up calls.

For repeated use, install the command on your PATH:

```powershell
npm install --global @sbroenne/excelcli
excelcli --version
```

The launcher installs `@sbroenne/excelcli-win32-x64` or
`@sbroenne/excelcli-win32-arm64` as an optional dependency, matching the Node.js
process architecture. x64 Node.js on ARM64 Windows still uses x64 emulation.
There is no automatic fallback if the matching runtime is missing.
Do not use `--omit=optional`. It forwards arguments,
standard input/output, and exit codes to the matching native executable; session
management and Excel behavior are unchanged.
Apple Silicon macOS selects `@sbroenne/excelcli-darwin-arm64`.

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
mkdir -p "$HOME/.local/bin/excelmcp"
unzip ExcelMcp-CLI-1.x.x-macos-arm64.zip -d "$HOME/.local/bin/excelmcp"
chmod +x "$HOME/.local/bin/excelmcp/excelcli"
```

### Add CLI to PATH

**Windows:**

```powershell
$toolsDir = "C:\Tools\ExcelMcp"
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*$toolsDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$userPath;$toolsDir", "User")
    Write-Host "Added $toolsDir to user PATH. Restart your terminal to apply."
}
```

Or manually: **Settings → System → About → Advanced system settings → Environment Variables → User variables → Path → Edit → New** → add `C:\Tools\ExcelMcp`

**macOS:** Add `export PATH="$HOME/.local/bin/excelmcp:$PATH"` to your shell
profile (for example `~/.zprofile` for zsh), then open a new terminal.
Use `command -v excelcli` to confirm which executable resolves.

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

**After installation:** Plugin guidance uses `npx -y @sbroenne/excelcli@latest`;
Node.js 18+ is required, but a global CLI installation is optional. Install the
global npm package or standalone executable above only if you want `excelcli`
on PATH for other scripts. The plugin's argument-safe PowerShell wrapper
requires PowerShell 7 and does not install a global shim or change PATH.

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

If the launcher cannot find the matching Windows or Darwin ARM64 runtime, reinstall with
optional dependencies enabled:

```powershell
npm install --global @sbroenne/excelcli --include=optional
```

The npm launcher supports Apple Silicon macOS. It reports an explicit error on
Linux, Intel Macs, and unsupported Windows architectures rather than selecting
an incompatible runtime.

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
# Solution: Install desktop Excel 2016+ on Windows or Excel for Mac 16.112+
```

### VBA Access Denied

**Windows only.** VBA is unavailable in the Mac beta. Reading/changing VBA
modules requires **"Trust access to the VBA project object model"** manually:

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

**Solution:** Reconcile the target workbook if it is open or already owned by a
session. Do not close unrelated Excel windows. On Mac, `RecoveryRequired` means
an uncertain handoff needs manual reconciliation before another attempt.

### Permission Issues

Use the returned access/permission error to identify the denied operation;
do not elevate privileges as a general workaround. Mac Automation permission
is granted manually for the requesting process identity. See
[platform troubleshooting](https://excelmcpserver.dev/troubleshooting/).

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
   - [Agent Skills](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli/skills/excel-cli) - token-efficient AI guidance for coding agents
4. **Join the community:** Star the repo, report issues, contribute improvements

**Happy automating! 🚀**

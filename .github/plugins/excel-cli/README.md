# Excel CLI Plugin

Source-owned README overlay copied into the published `excel-cli` GitHub Copilot plugin bundle. Install the plugin from the published marketplace repo, not from `.github\plugins\excel-cli`.

## Prerequisites

- **Windows** (COM interop required — macOS/Linux unsupported)
- **Microsoft Excel 2016 or later**

## Installation

### Step 1: Install the Plugin

```powershell
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install the CLI plugin
copilot plugin install excel-cli@mcp-server-excel-plugins
```

### Step 2: Wire the Bundled CLI onto PATH

The plugin now ships the actual `excelcli.exe` binary in `bin\`. Run the helper once to expose it as the `excelcli` command from any terminal:

```powershell
pwsh -ExecutionPolicy Bypass -File "$env:USERPROFILE\.copilot\installed-plugins\mcp-server-excel-plugins\excel-cli\bin\install-global.ps1"
```

The helper:
- creates `~/.copilot/bin` if needed
- writes `excelcli.cmd` and `excelcli.ps1` shims that point at the plugin-bundled binary
- adds `~/.copilot/bin` to your user PATH if missing

### Step 3: Verify

```powershell
excelcli --version
excelcli --help
```

## What's Included

- **Bundled `excelcli.exe`** — self-contained Windows CLI, no separate download
- **`excel-cli` skill** — token-efficient Excel automation guidance for coding agents
- **Install helper** — one-time PATH wiring for the bundled CLI

## Notes

- If you reinstall the plugin or test a locally built plugin path, re-run `install-global.ps1` so the shim points at the current plugin directory.
- The standalone ZIP and NuGet tool remain available for non-plugin installs, but they are no longer required for the Copilot plugin path.

## Support

Report issues at [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel/issues)

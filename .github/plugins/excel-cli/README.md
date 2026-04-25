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

### Step 2: Install `excelcli`

Install the CLI separately, then keep using the plugin for guidance:

**Option A: Standalone ZIP**
1. Download `ExcelMcp-CLI-{version}-windows.zip` from [Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract `excelcli.exe` to a permanent folder
3. Add that folder to your PATH

**Option B: .NET Tool**
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

### Step 3: Verify

```powershell
excelcli --version
excelcli --help
```

## What's Included

- **`excel-cli` skill** — token-efficient Excel automation guidance for coding agents

## Notes

- The plugin is skill-only; install `excelcli` separately via the standalone ZIP or NuGet tool.

## Support

Report issues at [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel/issues)

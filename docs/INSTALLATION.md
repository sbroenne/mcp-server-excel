# ExcelMcp CLI Installation Guide

Complete installation guide for the ExcelMcp CLI tool for direct Excel automation and development workflows.

## 🎯 System Requirements

- **Windows OS** - Required for Excel COM interop
- **Microsoft Excel** - Must be installed on the machine (2016+)
- **.NET 8 Runtime** - Install via: `winget install Microsoft.DotNet.Runtime.8`

> **Note:** For the MCP Server (AI assistant integration), see the [main README](../README.md).

---

## 🔧 CLI Installation

### Option 1: Download Binary (Recommended)

1. **Download the latest CLI release**:
   - Go to [Releases](https://github.com/sbroenne/mcp-server-excel/releases)
   - Download `ExcelMcp-CLI-{version}-windows.zip`

2. **Extract and install**:

   ```powershell
   # Extract to your preferred location
   Expand-Archive -Path "ExcelMcp-CLI-1.0.0-windows.zip" -DestinationPath "C:\Tools\ExcelMcp-CLI"
   
   # Add to PATH (optional but recommended)
   $env:PATH += ";C:\Tools\ExcelMcp-CLI"
   
   # Make PATH change permanent
   [Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
   ```

3. **Verify installation**:

   ```powershell
   # Check CLI version and help
   excelcli.exe
   
   # Test CLI functionality
   excelcli.exe create-empty "test.xlsx"
   ```

### CLI Quick Start

```powershell
# Basic operations
excelcli.exe create-empty "workbook.xlsx"
excelcli.exe pq-list "workbook.xlsx"
excelcli.exe sheet-read "workbook.xlsx" "Sheet1" "A1:D10"

# Power Query with privacy levels
excelcli.exe pq-import "workbook.xlsx" "MyQuery" "query.pq" --privacy-level Private

# VBA operations (requires one-time manual setup in Excel)
# Enable VBA trust: Excel → File → Options → Trust Center → Trust Center Settings
# → Macro Settings → Check "Trust access to the VBA project object model"
excelcli.exe create-empty "macros.xlsm"
excelcli.exe script-list "macros.xlsm"
```

---

## 🔨 Build from Source

### Prerequisites

- Windows OS with Excel installed
- .NET 8 SDK ([Download](https://dotnet.microsoft.com/download/dotnet/8.0))
- Git (for cloning the repository)

### Build Steps

1. **Clone the repository**:

   ```powershell
   git clone https://github.com/sbroenne/mcp-server-excel.git
   cd mcp-server-excel
   ```

2. **Build the solution**:

   ```powershell
   dotnet build -c Release
   ```

3. **Run tests** (requires Excel installed locally):

   ```powershell
   # Run unit tests only (no Excel required)
   dotnet test --filter "Category=Unit"
   
   # Run integration tests (requires Excel)
   dotnet test --filter "Category=Integration"
   ```

### After Building

**CLI Tool:**

```powershell
# CLI executable location
.\src\ExcelMcp.CLI\bin\Release\net10.0\excelcli.exe

# Add to PATH for easier access
$buildPath = "$(Get-Location)\src\ExcelMcp.CLI\bin\Release\net10.0"
$env:PATH += ";$buildPath"
[Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")

# Test CLI
excelcli.exe create-empty "test.xlsx"
```

### Installation Options

#### Option 1: Add to PATH (Recommended)

```powershell
# Add the build directory to your system PATH
$buildPath = "$(Get-Location)\src\ExcelMcp.CLI\bin\Release\net10.0"
$env:PATH += ";$buildPath"

# Make permanent (requires admin privileges)
[Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
```

#### Option 2: Copy to a tools directory

```powershell
# Create tools directory
New-Item -ItemType Directory -Path "C:\Tools\ExcelMcp-CLI" -Force

# Copy CLI files
Copy-Item "src\ExcelMcp.CLI\bin\Release\net10.0\*" "C:\Tools\ExcelMcp-CLI\" -Recurse

# Add to PATH
$env:PATH += ";C:\Tools\ExcelMcp-CLI"
[Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
```

---

## 🔧 VBA Configuration

### Required for VBA script operations

If you plan to use VBA script commands, you must manually enable VBA trust in Excel (one-time setup):

**Steps to Enable VBA Trust:**

1. Open Microsoft Excel
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice to save settings

After enabling this setting, VBA operations will work automatically. If VBA trust is not enabled, commands will display these instructions.

**Security Note:** ExcelMcp never modifies registry settings or security configurations automatically. Users must explicitly enable VBA trust through Excel's settings to maintain security control.

---

## 🔧 Power Query Privacy Configuration

### Optional: Set default privacy level for automation

For automated workflows, you can set a default privacy level via environment variable:

```powershell
# Set default privacy level (None, Private, Organizational, or Public)
$env:EXCEL_DEFAULT_PRIVACY_LEVEL = "Private"

# Make it permanent (optional)
[Environment]::SetEnvironmentVariable("EXCEL_DEFAULT_PRIVACY_LEVEL", "Private", "User")
```

Without this setting, commands will prompt for privacy level when needed, providing recommendations based on your existing queries.

---

## 🆘 Troubleshooting

### Common Issues

**"Excel is not installed" error:**

- Ensure Microsoft Excel is installed and accessible
- Try running Excel manually first to ensure it works

**"COM interop failed" error:**

- Restart your computer after Excel installation
- Check that Excel is not running with administrator privileges while your tool runs without

**".NET runtime not found" error:**

- Install .NET 8 Runtime: `winget install Microsoft.DotNet.Runtime.8`
- Verify installation: `dotnet --version`

**VBA access denied:**

- Enable VBA trust manually in Excel: File → Options → Trust Center → Trust Center Settings → Macro Settings → Check "Trust access to the VBA project object model"
- Restart Excel after enabling the setting
- If VBA commands still fail, ensure Excel is not running with elevated privileges

### Getting Help

- 📖 **Documentation**: [Complete command reference](COMMANDS.md)
- 🧠 **MCP Server Guide**: [MCP Server README](../src/ExcelMcp.McpServer/README.md)
- 🔧 **CLI Guide**: [CLI documentation](CLI.md)
- 🐛 **Issues**: [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)

---

## 📋 CLI Command Summary

| Category | Commands | Description |
|----------|----------|-------------|
| **File Operations** | `create-empty` | Create Excel workbooks (.xlsx, .xlsm) |
| **Power Query** | `pq-list`, `pq-view`, `pq-import`, `pq-export`, `pq-update`, `pq-refresh`, `pq-loadto`, `pq-delete` | Manage Power Query M code with optional `--privacy-level` parameter |
| **Worksheets** | `sheet-list`, `sheet-read`, `sheet-write`, `sheet-create`, `sheet-rename`, `sheet-copy`, `sheet-delete`, `sheet-clear`, `sheet-append` | Worksheet operations |
| **Parameters** | `param-list`, `param-get`, `param-set`, `param-create`, `param-delete` | Named range management |
| **Cells** | `cell-get-value`, `cell-set-value`, `cell-get-formula`, `cell-set-formula` | Individual cell operations |
| **VBA Scripts** | `script-list`, `script-export`, `script-import`, `script-update`, `script-run`, `script-delete` | VBA macro management (requires manual VBA trust setup in Excel) |

> **📋 [Complete Command Reference →](COMMANDS.md)** - Detailed documentation for all 40+ commands

---

## 🔄 Development & Contributing

**Important:** All changes to this project must be made through **Pull Requests (PRs)**. Direct commits to `main` are not allowed.

- 📋 **Development Workflow**: See [DEVELOPMENT.md](DEVELOPMENT.md) for complete PR process
- 🤝 **Contributing Guide**: See [CONTRIBUTING.md](CONTRIBUTING.md) for code standards

Version numbers are automatically managed by the release workflow - **do not update version numbers manually**.

- 📋 **Development Workflow**: See [DEVELOPMENT.md](DEVELOPMENT.md) for complete PR process
- 🤝 **Contributing Guide**: See [CONTRIBUTING.md](CONTRIBUTING.md) for code standards
- � **Release Strategy**: See [RELEASE-STRATEGY.md](RELEASE-STRATEGY.md) for release processes

Version numbers are automatically managed by the release workflow - **do not update version numbers manually**.

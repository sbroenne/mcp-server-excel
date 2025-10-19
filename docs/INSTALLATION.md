# Installation Guide

Choose your installation method based on your use case:

- **🧠 MCP Server**: For AI assistant integration (GitHub Copilot, Claude, ChatGPT)
- **🔧 CLI Tool**: For direct automation and development workflows
- **📦 Combined**: Both tools for complete Excel development environment

## 🎯 System Requirements

- **Windows OS** - Required for Excel COM interop
- **Microsoft Excel** - Must be installed on the machine (2016+)
- **.NET 10 SDK** - Install via: `winget install Microsoft.DotNet.SDK.10`

---

## 🧠 MCP Server Installation

**For AI assistant integration and conversational Excel workflows**

### Option 1: Microsoft's NuGet MCP Approach (Recommended)

Use the official `dnx` command to download and execute the MCP Server:

```powershell
# Download and execute MCP server using dnx
dnx Sbroenne.ExcelMcp.McpServer@latest --yes

# Execute specific version
dnx Sbroenne.ExcelMcp.McpServer@1.0.0 --yes

# Use with private feed
dnx Sbroenne.ExcelMcp.McpServer@latest --source https://your-feed.com --yes
```

**Benefits:**
- ✅ Official Microsoft approach for NuGet MCP servers
- ✅ Automatic download and execution in one command
- ✅ No separate installation step required
- ✅ Perfect for AI assistant integration
- ✅ Follows [Microsoft's NuGet MCP guidance](https://learn.microsoft.com/en-us/nuget/concepts/nuget-mcp)

### Option 2: Download Binary

1. **Download the latest MCP Server release**:
   - Go to [Releases](https://github.com/sbroenne/mcp-server-excel/releases)
   - Download `ExcelMcp-MCP-Server-{version}-windows.zip`

2. **Extract and run**:

   ```powershell
   # Extract to your preferred location
   Expand-Archive -Path "ExcelMcp-MCP-Server-1.0.0-windows.zip" -DestinationPath "C:\Tools\ExcelMcp-MCP"
   
   # Run the MCP server
   dotnet C:\Tools\ExcelMcp-MCP\ExcelMcp.McpServer.dll
   ```

### Configure with AI Assistants

**GitHub Copilot Integration:**

Add to your VS Code settings.json or MCP client configuration:

```json
{
  "mcp": {
    "servers": {
      "excel": {
        "command": "dnx",
        "args": ["Sbroenne.ExcelMcp.McpServer@latest", "--yes"],
        "description": "Excel development operations through MCP"
      }
    }
  }
}
```

**Claude Desktop Integration:**

Add to Claude Desktop MCP configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer@latest", "--yes"]
    }
  }
}
```

---

## 🔧 CLI Tool Installation

**For direct automation, development workflows, and CI/CD integration**

### Option 1: Download Binary (Recommended)

1. **Download the latest CLI release**:
   - Go to [Releases](https://github.com/sbroenne/mcp-server-excel/releases)
   - Download `ExcelMcp-CLI-{version}-windows.zip`

2. **Extract and install**:

   ```powershell
   # Extract to your preferred location
   Expand-Archive -Path "ExcelMcp-CLI-2.0.0-windows.zip" -DestinationPath "C:\Tools\ExcelMcp-CLI"
   
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

# VBA operations (requires one-time setup)
excelcli.exe setup-vba-trust
excelcli.exe create-empty "macros.xlsm"
excelcli.exe script-list "macros.xlsm"
```

---

## 📦 Combined Installation

**For users who need both MCP Server and CLI tools**

### Download Combined Package

1. **Download the combined release**:
   - Go to [Releases](https://github.com/sbroenne/mcp-server-excel/releases)
   - Download `ExcelMcp-{version}-windows.zip` (combined package)

2. **Extract and setup**:

   ```powershell
   # Extract to your preferred location
   Expand-Archive -Path "ExcelMcp-3.0.0-windows.zip" -DestinationPath "C:\Tools\ExcelMcp"
   
   # Add CLI to PATH
   $env:PATH += ";C:\Tools\ExcelMcp\CLI"
   [Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
   
   # Install MCP Server as .NET tool (from extracted files)
   dotnet tool install --global --add-source C:\Tools\ExcelMcp\MCP-Server ExcelMcp.McpServer
   ```

3. **Verify both tools**:

   ```powershell
   # Test CLI
   excelcli.exe create-empty "test.xlsx"
   
   # Test MCP Server
   mcp-excel --help
   ```

---

## 🔨 Build from Source

**For developers who want to build both tools from source**

### Prerequisites

- Windows OS with Excel installed
- .NET 10 SDK ([Download](https://dotnet.microsoft.com/download/dotnet/10.0))
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

**MCP Server:**

```powershell
# Run MCP server from build
dotnet run --project src/ExcelMcp.McpServer

# Or install as .NET tool from local build
dotnet pack src/ExcelMcp.McpServer -c Release
dotnet tool install --global --add-source src/ExcelMcp.McpServer/bin/Release ExcelMcp.McpServer
```

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

#### Option 1: Add to PATH (Recommended for coding agents)

```powershell
# Add the build directory to your system PATH
$buildPath = "$(Get-Location)\src\\ExcelMcp.CLI\\bin\Release\net10.0"
$env:PATH += ";$buildPath"

# Make permanent (requires admin privileges)
[Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
```

#### Option 2: Copy to a tools directory

```powershell
---

## 🔧 VBA Configuration

**Required for VBA script operations (both MCP Server and CLI)**

If you plan to use VBA script commands, configure VBA trust:

```powershell
# One-time setup for VBA automation (works with both tools)
# For CLI:
excelcli.exe setup-vba-trust

# For MCP Server (through AI assistant):
# Ask your AI assistant: "Setup VBA trust for Excel automation"
```

This configures the necessary registry settings to allow programmatic access to VBA projects.

---

## 📋 Installation Summary

| Use Case | Tool | Installation Method | Command |
|----------|------|-------------------|---------|
| **AI Assistant Integration** | MCP Server | .NET Tool | `dotnet tool install --global Sbroenne.ExcelMcp.McpServer` |
| **Direct Automation** | CLI | Binary Download | Download `ExcelMcp-CLI-{version}-windows.zip` |
| **Development/Testing** | Both | Build from Source | `git clone` + `dotnet build` |
| **Complete Environment** | Combined | Binary Download | Download `ExcelMcp-{version}-windows.zip` |

## 🆘 Troubleshooting

### Common Issues

**"Excel is not installed" error:**

- Ensure Microsoft Excel is installed and accessible
- Try running Excel manually first to ensure it works

**"COM interop failed" error:**

- Restart your computer after Excel installation
- Check that Excel is not running with administrator privileges while your tool runs without

**".NET runtime not found" error:**

- Install .NET 10 SDK: `winget install Microsoft.DotNet.SDK.10`
- Verify installation: `dotnet --version`

**VBA access denied:**

- Run the VBA trust setup command once
- Restart Excel after running the trust setup

### Getting Help

- 📖 **Documentation**: [Complete command reference](COMMANDS.md)
- 🧠 **MCP Server Guide**: [MCP Server README](../src/ExcelMcp.McpServer/README.md)
- 🔧 **CLI Guide**: [CLI documentation](CLI.md)
- 🐛 **Issues**: [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)

---

## 🔄 Development & Contributing

**Important:** All changes to this project must be made through **Pull Requests (PRs)**. Direct commits to `main` are not allowed.

- 📋 **Development Workflow**: See [DEVELOPMENT.md](DEVELOPMENT.md) for complete PR process
- 🤝 **Contributing Guide**: See [CONTRIBUTING.md](CONTRIBUTING.md) for code standards
- � **Release Strategy**: See [RELEASE-STRATEGY.md](RELEASE-STRATEGY.md) for release processes

Version numbers are automatically managed by the release workflow - **do not update version numbers manually**.

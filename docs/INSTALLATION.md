# Installation Guide

## Requirements

- **Windows OS** - Required for Excel COM interop
- **Microsoft Excel** - Must be installed on the machine
- **.NET 8.0 Runtime** - Download from [Microsoft](https://dotnet.microsoft.com/download/dotnet/8.0)

## 📦 Installation Options

### Option 1: Install MCP Server via NuGet (Recommended for MCP)

Install the MCP Server as a global .NET tool:

```powershell
# Install globally
dotnet tool install --global ExcelMcp.McpServer

# Run the MCP server
mcp-excel

# Update to latest version
dotnet tool update --global ExcelMcp.McpServer

# Uninstall
dotnet tool uninstall --global ExcelMcp.McpServer
```

**Benefits:**
- ✅ Easy installation with a single command
- ✅ Automatic updates via `dotnet tool update`
- ✅ Global availability from any directory
- ✅ Perfect for AI assistant integration

### Option 2: Download Pre-built Binary (Recommended for CLI)

You can download the latest pre-built version:

1. **Download the latest release**:
   - Go to [Releases](https://github.com/sbroenne/mcp-server-excel/releases)
   - Download `ExcelCLI-1.0.3-windows.zip` (or the latest version)

2. **Extract and install**:

   ```powershell
   # Extract to your preferred location
   Expand-Archive -Path "ExcelCLI-1.0.3-windows.zip" -DestinationPath "C:\Tools\ExcelCLI"
   
   # Add to PATH (optional but recommended)
   $env:PATH += ";C:\Tools\ExcelMcp\CLI"
   
   # Make PATH change permanent
   [Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
   ```

3. **Verify installation**:

   ```powershell
   # Check CLI version
   ExcelMcp --version
   
   # Test CLI functionality
   ExcelMcp create-empty "test.xlsx"
   
   # Check MCP Server (if downloaded binary)
   dotnet C:\Tools\ExcelMcp\MCP-Server\ExcelMcp.McpServer.dll
   ```

### Option 3: Build from Source

### Prerequisites for Building

- Windows OS with Excel installed
- .NET 8.0 SDK ([Download](https://dotnet.microsoft.com/download/dotnet/8.0))
- Git (for cloning the repository)

### Build Steps

1. **Clone the repository**:

   ```powershell
   git clone https://github.com/sbroenne/mcp-server-excel.git
   cd ExcelMcp
   ```

2. **Build the solution**:

   ```powershell
   dotnet build -c Release
   ```

3. **Run tests** (requires Excel installed locally):

   ```powershell
   dotnet test
   ```
   
   > **Note**: Tests require Microsoft Excel to be installed and accessible via COM automation. The GitHub Actions CI only builds the project (tests are skipped) since runners don't have Excel installed.

4. **Locate the executable**:

   ```text
   src\ExcelMcp\bin\Release\net8.0\ExcelMcp.exe
   ```

### Installation Options

#### Option 1: Add to PATH (Recommended for coding agents)

```powershell
# Add the build directory to your system PATH
$buildPath = "$(Get-Location)\src\ExcelMcp\bin\Release\net8.0"
$env:PATH += ";$buildPath"

# Make permanent (requires admin privileges)
[Environment]::SetEnvironmentVariable("PATH", $env:PATH, "User")
```

#### Option 2: Copy to a tools directory

```powershell
# Create a tools directory
mkdir C:\Tools\ExcelMcp
copy "src\ExcelMcp\bin\Release\net8.0\*" "C:\Tools\ExcelMcp\"

# Add to PATH
$env:PATH += ";C:\Tools\ExcelCLI"
```

#### Option 3: Use directly with full path

```powershell
# Use the full path in scripts
.\src\ExcelMcp\bin\Release\net8.0\ExcelMcp.exe pq-list "myfile.xlsx"
```

### Verify Installation

Test that ExcelMcp is working correctly:

```powershell
# Check version
ExcelMcp --version

# Test functionality
ExcelMcp create-empty "test.xlsx"
```

If successful, you should see confirmation that the Excel file was created.

## For VBA Operations

If you plan to use VBA script commands, you'll need to configure VBA trust:

```powershell
# One-time setup for VBA automation
ExcelMcp setup-vba-trust
```

This configures the necessary registry settings to allow programmatic access to VBA projects.

## 🔄 **Development & Contributing**

**Important:** All changes to this project must be made through **Pull Requests (PRs)**. Direct commits to `main` are not allowed.

- 📋 **Development Workflow**: See [DEVELOPMENT.md](DEVELOPMENT.md) for complete PR process
- 🤝 **Contributing Guide**: See [CONTRIBUTING.md](CONTRIBUTING.md) for code standards
- 🐛 **Report Issues**: Use [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues) for bugs and feature requests

Version numbers are automatically managed by the release workflow - **do not update version numbers manually**.

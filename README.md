# ExcelMcp - Excel MCP Server for AI-Powered Development

[![Build](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total)](https://github.com/sbroenne/mcp-server-excel/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

`MCP Server` • `Excel Development` • `AI Assistant Integration` • `Model Context Protocol` • `GitHub Copilot`

A **Model Context Protocol (MCP) server** that enables **AI assistants** like GitHub Copilot, Claude, and ChatGPT to perform Excel development tasks through conversational interfaces. Transform Excel development workflows with AI-assisted Power Query refactoring, VBA enhancement, code review, and automation—all through natural language interactions.

**🎯 How does it work?** The MCP server provides AI assistants with structured access to Excel operations through 6 resource-based tools, enabling conversational Excel development workflows while maintaining the power of direct Excel COM interop integration.

## 🤖 **AI-Powered Excel Development**

- **GitHub Copilot** - Native MCP integration for conversational Excel development
- **Claude & ChatGPT** - Natural language Excel operations through MCP protocol
- **Power Query AI Enhancement** - AI-assisted M code refactoring and optimization
- **VBA AI Development** - Intelligent macro enhancement with error handling
- **Conversational Workflows** - Ask AI to perform complex Excel tasks naturally

## 🚀 Quick Start

### Install MCP Server as .NET Tool (Recommended)

```powershell
# Install .NET 10 SDK first
winget install Microsoft.DotNet.SDK.10

# Install globally as .NET tool
dotnet tool install --global Sbroenne.ExcelMcp.McpServer

# Run MCP server
mcp-excel

# Update to latest version
dotnet tool update --global Sbroenne.ExcelMcp.McpServer

# Uninstall
dotnet tool uninstall --global Sbroenne.ExcelMcp.McpServer
```

### Configure with AI Assistants

**GitHub Copilot Integration:**

```json
// Add to your VS Code settings.json or MCP client configuration
{
  "mcp": {
    "servers": {
      "excel": {
        "command": "mcp-excel",
        "description": "Excel development operations through MCP"
      }
    }
  }
}
```

**Claude Desktop Integration:**

```json
// Add to Claude Desktop MCP configuration
{
  "mcpServers": {
    "excel": {
      "command": "mcp-excel",
      "args": []
    }
  }
}
```

### Build from Source

```powershell
# Clone and build
git clone https://github.com/sbroenne/mcp-server-excel.git
cd mcp-server-excel
dotnet build -c Release

# Run MCP server
dotnet run --project src/ExcelMcp.McpServer

# Run tests (requires Excel installed locally)
dotnet test --filter "Category=Unit"
```

## ✨ Key Features

- 🤖 **MCP Protocol Integration** - Native support for AI assistants through Model Context Protocol
- �️ **Conversational Interface** - Natural language Excel operations with AI assistants
- � **Resource-Based Architecture** - 6 structured tools instead of 40+ individual commands
- � **AI-Assisted Development** - Power Query refactoring, VBA enhancement, code review
- 🧠 **Smart Context Management** - AI assistants understand Excel development workflows
- 💾 **Full Excel Integration** - Controls actual Excel application for complete feature access
- �️ **Production Ready** - Enterprise-grade with security validation and robust error handling
- 📈 **Development Focus** - Optimized for Excel development, not data processing workflows

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Setup, configuration, and AI integration examples

## 📖 Documentation

| Document | Description |
|----------|-------------|
| **[🧠 MCP Server Guide](src/ExcelMcp.McpServer/README.md)** | Complete MCP server setup and AI integration examples |
| **[🔧 ExcelMcp.CLI](docs/CLI.md)** | Command-line interface for direct Excel automation |
| **[📋 Command Reference](docs/COMMANDS.md)** | Complete reference for all 40+ CLI commands |
| **[⚙️ Installation Guide](docs/INSTALLATION.md)** | Building from source and installation options |
| **[🤖 GitHub Copilot Integration](docs/COPILOT.md)** | Using ExcelMcp with GitHub Copilot |
| **[🔧 Development Workflow](docs/DEVELOPMENT.md)** | Contributing guidelines and PR requirements |
| **[📦 NuGet Publishing](docs/NUGET_TRUSTED_PUBLISHING.md)** | Trusted publishing setup for maintainers |

## 🎯 Excel Development Use Cases

### **Power Query Development**

- **M Code Refactoring** - AI-assisted optimization of Power Query transformations
- **Performance Analysis** - Identify and fix slow Power Query operations
- **Code Review** - Analyze M code for best practices and improvements
- **Version Control** - Export/import Power Query code for Git workflows

### **VBA Development & Enhancement**

- **Error Handling** - Add comprehensive try-catch patterns to VBA modules
- **Code Quality** - Implement logging, input validation, and best practices
- **Module Management** - Export, enhance, and import VBA code with AI assistance
- **Testing & Debugging** - Automated testing workflows for Excel macros

### **Excel Development Workflows**

- **CI/CD Integration** - Automated Excel development testing and validation
- **Code Templates** - Generate Excel workbook templates for development projects
- **Development Environment** - Create and configure Excel files for coding workflows
- **Documentation** - Generate code documentation and comments for Excel automation

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server | COM interop is Windows-specific |
| **Microsoft Excel** | Any recent version (2016+) | ExcelMcp controls the actual Excel application |
| **.NET 8.0 Runtime** | [Download here](https://dotnet.microsoft.com/download/dotnet/8.0) | ExcelMcp runtime dependency |

> **🚨 Critical:** ExcelMcp controls the actual running Excel application through COM interop, not just Excel file formats. This provides access to Excel's full feature set (Power Query engine, VBA runtime, formula calculations, charts, pivot tables) but requires Excel to be installed and available for automation.

## 6️⃣ MCP Tools Overview

The MCP server provides 6 resource-based tools for AI assistants:

- **excel_file** - File management (create, validate, check-exists)
- **excel_powerquery** - Power Query operations (list, view, import, export, update, refresh, delete)
- **excel_worksheet** - Worksheet operations (list, read, write, create, rename, copy, delete, clear, append)
- **excel_parameter** - Named range management (list, get, set, create, delete)
- **excel_cell** - Cell operations (get-value, set-value, get-formula, set-formula)
- **excel_vba** - VBA script management (list, export, import, update, run, delete)

> 🧠 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Detailed MCP integration and AI examples

## 🔗 Additional Tools

- **[ExcelMcp.CLI](docs/CLI.md)** - Command-line interface for direct Excel automation
- **[Command Reference](docs/COMMANDS.md)** - All 40+ CLI commands for script-based workflows

## 🤝 Contributing

We welcome contributions! See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏷️ **SEO & Discovery Tags**

**MCP Server** | **Model Context Protocol** | **AI Excel Development** | **GitHub Copilot Excel** | **Excel MCP** | **Conversational Excel** | **AI Assistant Integration** | **Power Query AI** | **VBA AI Development** | **Excel Code Review** | **Excel COM Interop** | **Excel Development AI**

## �🙏 Acknowledgments

- **GitHub Copilot** - This entire project was developed using AI assistance
- **Microsoft Excel Team** - For the comprehensive COM automation APIs
- **Open Source Community** - For inspiration and best practices in CLI tool development

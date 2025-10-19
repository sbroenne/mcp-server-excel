# ExcelMcp - Excel MCP Server for AI-Powered Development

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
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

### Install & Configure (2 Steps)

#### Step 1: Install .NET 10 SDK

```powershell
winget install Microsoft.DotNet.SDK.10
```

#### Step 2: Configure GitHub Copilot MCP Server

Create or modify `.vscode/mcp.json` in your workspace:

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

That's it! The `dnx` command automatically downloads and runs the latest version when GitHub Copilot needs it.

## 🧠 **GitHub Copilot Integration**

To make GitHub Copilot aware of ExcelMcp in your own projects:

1. **Copy the Copilot Instructions** to your project:

   ```bash
   # Copy ExcelMcp automation instructions to your project's .github directory
   curl -o .github/copilot-instructions.md https://raw.githubusercontent.com/sbroenne/mcp-server-excel/main/docs/excel-powerquery-vba-copilot-instructions.md
   ```

2. **Configure VS Code** (optional but recommended):

   ```json
   {
     "github.copilot.enable": {
       "*": true,
       "csharp": true,
       "powershell": true,
       "yaml": true
     }
   }
   ```

### **Effective Copilot Prompting**

With the ExcelMcp instructions installed, Copilot will automatically suggest Excel operations through the MCP server. Here's how to get the best results:

```text
"Use the excel MCP server to..." - Reference the configured server name
"Create an Excel workbook with Power Query that..." - Natural language Excel tasks
"Help me debug this Excel automation issue..." - For troubleshooting assistance
"Export the VBA code from this Excel file..." - Specific Excel operations
```

### Alternative AI Assistants

**Claude Desktop Integration:**

Add to your Claude Desktop MCP configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer", "--yes"]
    }
  }
}
```

**Direct Command Line Testing:**

```powershell
# Test the MCP server directly
dnx Sbroenne.ExcelMcp.McpServer --yes
```

> **Note:** The MCP server will appear to "hang" after startup - this is expected behavior as it waits for MCP protocol messages from your AI assistant.

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
| **.NET 10 SDK** | `winget install Microsoft.DotNet.SDK.10` | Required for `dnx` command execution |

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

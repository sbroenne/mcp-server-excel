# ExcelMcp - Excel MCP Server and CLI for AI-Powered Development

[![Build](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![NuGet](https://img.shields.io/nuget/v/ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/ExcelMcp.McpServer)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total)](https://github.com/sbroenne/mcp-server-excel/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

`Excel Development` • `Power Query CLI` • `VBA Command Line` • `Excel COM Interop` • `MCP Server`

A comprehensive command-line interface tool designed specifically for **Excel development workflows** with **AI assistants** and **coding agents**. ExcelMcp enables GitHub Copilot, Claude, ChatGPT, and other AI tools to refactor Power Query M code, enhance VBA macros, review Excel automation code, and manage development workflows—all without requiring the Excel UI.

**🎯 How does it work?** ExcelMcp controls the actual Excel application itself (not just Excel files), providing access to Power Query transformations, VBA execution environment, and all native Excel development features through COM interop.

## 🔍 **Perfect for Excel Development with AI Assistants**

- **GitHub Copilot** - Built specifically for AI-powered Excel development workflows
- **Claude, ChatGPT, Cursor** - Command-line interface ideal for Excel code development
- **Power Query Development** - Refactor and optimize M code with AI assistance
- **VBA Development** - Enhance macros with error handling and best practices
- **Code Review & Testing** - Automated Excel development workflows and CI/CD

## 🚀 Quick Start

### Option 1: Download Pre-built Binary (Recommended)

```powershell
# Download from https://github.com/sbroenne/mcp-server-excel/releases
# Extract ExcelCLI-1.0.3-windows.zip to C:\Tools\ExcelMcp

# Add to PATH (optional)
$env:PATH += ";C:\Tools\ExcelCLI"

# Basic usage
ExcelMcp --version
ExcelMcp create-empty "test.xlsx"
ExcelMcp sheet-read "test.xlsx" "Sheet1"

# For VBA operations (one-time setup)
ExcelMcp setup-vba-trust
ExcelMcp create-empty "macros.xlsm"
```

### Option 2: Install MCP Server as .NET Tool (NuGet)

```powershell
# Install globally as .NET tool
dotnet tool install --global ExcelMcp.McpServer

# Run MCP server
mcp-excel

# Update to latest version
dotnet tool update --global ExcelMcp.McpServer

# Uninstall
dotnet tool uninstall --global ExcelMcp.McpServer
```

### Option 3: Build from Source

```powershell
# Clone and build
git clone https://github.com/sbroenne/mcp-server-excel.git
cd ExcelMcp
dotnet build -c Release

# Run tests (requires Excel installed locally)
dotnet test

# Basic usage
.\src\ExcelMcp\bin\Release\net8.0\ExcelMcp.exe --version
.\src\ExcelMcp\bin\Release\net8.0\ExcelMcp.exe create-empty "test.xlsx"
```

## ✨ Key Features

- 🤖 **AI Development Assistant** - Built for GitHub Copilot and AI-assisted Excel development
- 🔧 **Power Query Development** - Extract, refactor, and version control M code with AI assistance
- 📋 **VBA Development Tools** - Enhance macros, add error handling, and manage VBA modules
- 📊 **Excel Development API** - 40+ commands for Excel development workflows and testing
- 🛡️ **Production Ready** - Enterprise-grade with security validation and robust error handling
- 💾 **COM Interop Excellence** - Controls the actual Excel application for full development access
- 🔄 **Development Integration** - Perfect for CI/CD pipelines and Excel development workflows
- 📈 **Code Quality Focus** - Excel development with proper testing and code review practices

## 🧠 **MCP Server for AI Development** *(New!)*

ExcelMcp includes a **Model Context Protocol (MCP) server** for AI assistants like GitHub Copilot to provide conversational Excel development workflows - Power Query refactoring, VBA enhancement, and code review.

> 📚 **[Complete MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** - Setup, configuration, and AI integration examples

## 🏷️ **Keywords & Technologies**

**Primary:** `Excel CLI`, `Excel Development`, `Power Query CLI`, `VBA Command Line`, `Excel COM Interop`

**AI/Development:** `GitHub Copilot`, `MCP Server`, `Model Context Protocol`, `AI Assistant`, `Code Refactoring`, `VBA Development`

**Technologies:** `.NET 8`, `C#`, `COM Interop`, `Windows`, `PowerShell`, `Command Line Interface`, `MCP Protocol`

**Excel Features:** `Power Query M Code`, `VBA Macros`, `Excel Worksheets`, `Named Ranges`, `Excel Formulas`

**Use Cases:** `Excel Development`, `Power Query Refactoring`, `VBA Coding`, `Code Review`, `Development Testing`

## 📖 Documentation

| Document | Description |
|----------|-------------|
| **[📋 Command Reference](docs/COMMANDS.md)** | Complete reference for all 40+ ExcelMcp commands |
| **[🧠 MCP Server](src/ExcelMcp.McpServer/README.md)** | Model Context Protocol server for AI assistant integration |
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

## 🔄 **Compared to Excel Development Alternatives**

| Feature | ExcelMcp | Python openpyxl | Excel VBA | PowerShell Excel |
|---------|----------|-----------------|-----------|------------------|
| **Power Query Development** | ✅ Full M code access | ❌ No support | ❌ Limited | ❌ No support |
| **VBA Development** | ✅ Full module control | ❌ No support | ✅ Native | ✅ Limited |
| **AI Development Assistant** | ✅ Built for Copilot | ⚠️ Requires custom setup | ❌ Manual only | ⚠️ Complex integration |
| **Development Approach** | ✅ **Excel App Control** | ❌ File parsing only | ✅ **Excel App Control** | ✅ **Excel App Control** |
| **CLI Development Tools** | ✅ Purpose-built CLI | ⚠️ Script required | ❌ No CLI | ⚠️ Complex commands |
| **Excel Installation** | ⚠️ **Required** | ✅ Not needed | ⚠️ **Required** | ⚠️ **Required** |
| **Cross-Platform** | ❌ Windows only | ✅ Cross-platform | ❌ Windows only | ❌ Windows only |

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server | COM interop is Windows-specific |
| **Microsoft Excel** | Any recent version (2016+) | ExcelMcp controls the actual Excel application |
| **.NET 8.0 Runtime** | [Download here](https://dotnet.microsoft.com/download/dotnet/8.0) | ExcelMcp runtime dependency |

> **🚨 Critical:** ExcelMcp controls the actual running Excel application through COM interop, not just Excel file formats. This provides access to Excel's full feature set (Power Query engine, VBA runtime, formula calculations, charts, pivot tables) but requires Excel to be installed and available for automation.

## 📋 Commands Overview

ExcelMcp provides 40+ commands across 6 categories:

- **File Operations** (2) - Create Excel workbooks (.xlsx, .xlsm)
- **Power Query** (8) - M code management and data transformation  
- **VBA Scripts** (6) - Macro development and execution
- **Worksheets** (9) - Data manipulation and sheet management
- **Parameters** (5) - Named range configuration
- **Cells** (4) - Individual cell operations

> 📖 **[Complete Command Reference →](docs/COMMANDS.md)** - Detailed syntax and examples for all commands

## 🤝 Contributing

We welcome contributions! See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏷️ **SEO & Discovery Tags**

**Excel CLI** | **Power Query CLI** | **VBA Command Line** | **Excel Development** | **MCP Server** | **Model Context Protocol** | **AI Excel Development** | **GitHub Copilot Excel** | **Power Query Refactoring** | **VBA Development** | **Excel Code Review** | **Excel COM Interop** | **Excel DevOps**

## �🙏 Acknowledgments

- **GitHub Copilot** - This entire project was developed using AI assistance
- **Microsoft Excel Team** - For the comprehensive COM automation APIs
- **Open Source Community** - For inspiration and best practices in CLI tool development

# ExcelMcp.CLI - Excel Command Line Interface

[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)

`Excel Automation` • `Power Query CLI` • `VBA Command Line` • `Excel COM Interop` • `Development Tools`

A comprehensive command-line interface tool for **Excel development workflows** and **automation**. ExcelMcp.CLI provides direct programmatic access to Microsoft Excel through COM interop, enabling developers to manage Excel workbooks, Power Query transformations, VBA scripts, and data operations without requiring the Excel UI.

**🎯 How does it work?** ExcelMcp.CLI controls the actual Excel application itself (not just Excel files), providing access to Power Query transformations, VBA execution environment, and all native Excel development features through COM interop.

## 🚀 Quick Start

### Download Pre-built Binary (Recommended)

```powershell
# Download from https://github.com/sbroenne/mcp-server-excel/releases
# Extract excelcli-{version}-windows.zip to C:\Tools\excelcli

# Add to PATH (optional)
$env:PATH += ";C:\Tools\excelcli"

# Basic usage
excelcli create-empty "test.xlsx"
excelcli sheet-read "test.xlsx" "Sheet1"

# For VBA operations (one-time manual setup in Excel)
# Enable VBA trust: Excel → File → Options → Trust Center → Trust Center Settings
# → Macro Settings → Check "Trust access to the VBA project object model"
excelcli create-empty "macros.xlsm"
```

### Build from Source

```powershell
# Clone and build
git clone https://github.com/sbroenne/mcp-server-excel.git
cd mcp-server-excel
dotnet build -c Release

# Run tests (requires Excel installed locally)
dotnet test --filter "Category=Unit"

# Basic usage
.\src\ExcelMcp.CLI\bin\Release\net10.0\excelcli.exe create-empty "test.xlsx"
```

## ✨ Key Features

- 🔧 **Power Query Development** - Extract, refactor, and version control M code
- 📋 **VBA Development Tools** - Manage VBA modules, run macros, and enhance code
- 📊 **Excel Automation API** - 40+ commands for Excel operations and workflows
- 💾 **COM Interop Excellence** - Controls the actual Excel application for full access
- 🛡️ **Production Ready** - Enterprise-grade with security validation and robust error handling
- 🔄 **CI/CD Integration** - Perfect for automated Excel development workflows
- 📈 **Development Focus** - Excel development with proper testing and code practices

## 🎯 Excel Development Use Cases

### **Power Query Development**

- **M Code Management** - Export, import, and update Power Query transformations
- **Performance Testing** - Refresh and validate Power Query operations
- **Code Review** - Analyze M code for optimization opportunities
- **Version Control** - Export/import Power Query code for Git workflows

### **VBA Development & Enhancement**

- **Module Management** - Export, import, and update VBA modules
- **Macro Execution** - Run VBA macros with parameters from command line
- **Code Quality** - Implement error handling and best practices
- **Testing & Debugging** - Automated testing workflows for Excel macros

### **Excel Automation Workflows**

- **Data Processing** - Automate Excel data manipulation tasks
- **Report Generation** - Create and populate Excel reports programmatically
- **Configuration Management** - Manage Excel parameters through named ranges
- **Batch Operations** - Process multiple Excel files in automated workflows

## ⚙️ System Requirements

| Requirement | Details | Why Required |
|-------------|---------|--------------|
| **Windows OS** | Windows 10/11 or Server | COM interop is Windows-specific |
| **Microsoft Excel** | Any recent version (2016+) | ExcelMcp.CLI controls the actual Excel application |
| **.NET 10 Runtime** | [Download here](https://dotnet.microsoft.com/download/dotnet/10.0) | ExcelMcp.CLI runtime dependency |

> **🚨 Critical:** ExcelMcp.CLI controls the actual running Excel application through COM interop, not just Excel file formats. This provides access to Excel's full feature set (Power Query engine, VBA runtime, formula calculations) but requires Excel to be installed and available for automation.

## 📋 Commands Overview

ExcelMcp.CLI provides 40+ commands across 6 categories:

- **File Operations** (2) - Create Excel workbooks (.xlsx, .xlsm)
- **Power Query** (8) - M code management and data transformation  
- **VBA Scripts** (6) - Macro development and execution
- **Worksheets** (9) - Data manipulation and sheet management
- **Parameters** (5) - Named range configuration
- **Cells** (4) - Individual cell operations

> 📖 **[Complete Command Reference →](COMMANDS.md)** - Detailed syntax and examples for all commands

## 🔗 Related Tools

- **[ExcelMcp MCP Server](../README.md)** - Model Context Protocol server for AI assistant integration

## 📖 Documentation

| Document | Description |
|----------|-------------|
| **[📋 Command Reference](COMMANDS.md)** | Complete reference for all 40+ ExcelMcp.CLI commands |
| **[⚙️ Installation Guide](INSTALLATION.md)** | Building from source and installation options |
| **[🔧 Development Workflow](DEVELOPMENT.md)** | Contributing guidelines and PR requirements |

## 🤝 Contributing

We welcome contributions! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](../LICENSE) file for details.

---

**🏠 [← Back to Main Project](../README.md)** | **🧠 [MCP Server →](../src/ExcelMcp.McpServer/README.md)**

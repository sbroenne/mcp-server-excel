# ExcelMcp Model Context Protocol (MCP) Server

[![NuGet](https://img.shields.io/nuget/v/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Sbroenne.ExcelMcp.McpServer.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)
[![MCP Server](https://img.shields.io/badge/MCP%20Server-NuGet-blue.svg)](https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer)

The ExcelMcp MCP Server provides AI assistants with powerful Excel automation capabilities through the official Model Context Protocol (MCP) SDK. This enables natural language interactions with Excel through AI coding assistants like **GitHub Copilot**, **Claude**, and **ChatGPT** using a modern resource-based architecture.

## 🚀 Quick Start

### Option 1: Microsoft's NuGet MCP Approach (Recommended)

```bash
# Download and execute using dnx command
dnx Sbroenne.ExcelMcp.McpServer --yes
```

This follows Microsoft's official [NuGet MCP approach](https://learn.microsoft.com/en-us/nuget/concepts/nuget-mcp) where the `dnx` command automatically downloads and executes the MCP server from NuGet.org.

### Option 2: Build and Run from Source

```bash
# Build the MCP server
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# Run the MCP server (stdio mode)
dotnet run --project src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
```

### Configuration with AI Assistants

**For NuGet MCP Installation (dnx):**

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

**For Source Build:**

```json
{
  "servers": {
    "excel": {
      "command": "dotnet", 
      "args": ["run", "--project", "path/to/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj"]
    }
  }
}
```

## 🛠️ Resource-Based Tools

The MCP server provides **6 powerful resource-based tools** that follow REST-like design principles. Each tool supports multiple actions through a single, consistent interface:

### 1. **`excel_file`** - File Management

**Actions**: `create-empty`, `validate`, `check-exists`

- Create new Excel workbooks (.xlsx or .xlsm)
- Validate file format and existence
- Check file properties and status

### 2. **`excel_powerquery`** - Power Query Management  

**Actions**: `list`, `view`, `import`, `export`, `update`, `refresh`, `loadto`, `delete`

- Manage M code and data transformations
- Import/export queries for version control
- Refresh data connections and load to worksheets

### 3. **`excel_worksheet`** - Worksheet Operations

**Actions**: `list`, `read`, `write`, `create`, `rename`, `copy`, `delete`, `clear`, `append`  

- CRUD operations on worksheets and data ranges
- Bulk data import/export with CSV support
- Dynamic worksheet management

### 4. **`excel_parameter`** - Named Range Management

**Actions**: `list`, `get`, `set`, `create`, `delete`

- Manage named ranges as configuration parameters
- Get/set parameter values for dynamic workbooks
- Create and manage parameter schemas

### 5. **`excel_cell`** - Cell Operations

**Actions**: `get-value`, `set-value`, `get-formula`, `set-formula`

- Individual cell value and formula operations
- Precise cell-level data manipulation
- Formula validation and management

### 6. **`excel_vba`** - VBA Script Management ⚠️ *(.xlsm files only)*

**Actions**: `list`, `export`, `import`, `update`, `run`, `delete`, `setup-trust`, `check-trust`

- VBA module management and execution
- Script import/export for version control  
- Trust configuration for macro execution

## 💬 Example AI Assistant Interactions

### File Management

```text
User: "Create a new Excel workbook for my quarterly report"
AI Assistant uses: excel_file(action="create-empty", filePath="quarterly-report.xlsx")
Result: {"success": true, "filePath": "quarterly-report.xlsx", "message": "Excel file created successfully"}

User: "Check if my budget file exists and get its size"
AI Assistant uses: excel_file(action="check-exists", filePath="budget.xlsx")  
Result: {"exists": true, "filePath": "budget.xlsx", "size": 2048576}
```

### Power Query Management

```text
User: "Show me all Power Queries in my sales report"
AI Assistant uses: excel_powerquery(action="list", filePath="sales-report.xlsx")
Result: {"success": true, "action": "list", "filePath": "sales-report.xlsx"}

User: "Export the M code for the 'CustomerData' query"
AI Assistant uses: excel_powerquery(action="export", filePath="sales-report.xlsx", queryName="CustomerData", sourceOrTargetPath="customer-query.pq")
Result: {"success": true, "action": "export", "filePath": "sales-report.xlsx"}
```

### Worksheet Operations

```text
User: "List all worksheets in my analysis workbook"
AI Assistant uses: excel_worksheet(action="list", filePath="analysis.xlsx")
Result: {"success": true, "action": "list", "filePath": "analysis.xlsx"}

User: "Create a new worksheet called 'Summary'"
AI Assistant uses: excel_worksheet(action="create", filePath="analysis.xlsx", sheetName="Summary")
Result: {"success": true, "action": "create", "filePath": "analysis.xlsx"}
```

### Cell Operations

```text  
User: "What's the value in cell B5 of the Summary sheet?"
AI Assistant uses: excel_cell(action="get-value", filePath="report.xlsx", sheetName="Summary", cellAddress="B5")
Result: {"success": true, "action": "get-value", "filePath": "report.xlsx", "sheetName": "Summary", "cellAddress": "B5"}

User: "Set cell A1 to contain the formula =SUM(B1:B10)"  
AI Assistant uses: excel_cell(action="set-formula", filePath="report.xlsx", sheetName="Sheet1", cellAddress="A1", valueOrFormula="=SUM(B1:B10)")
Result: {"success": true, "action": "set-formula", "filePath": "report.xlsx", "sheetName": "Sheet1", "cellAddress": "A1"}
```

### Parameter Management

```text
User: "List all parameters in my configuration file"
AI Assistant uses: excel_parameter(action="list", filePath="config.xlsx")
Result: {"success": true, "action": "list", "filePath": "config.xlsx"}

User: "Set the StartDate parameter to 2024-01-01"
AI Assistant uses: excel_parameter(action="set", filePath="config.xlsx", paramName="StartDate", valueOrReference="2024-01-01")
Result: {"success": true, "action": "set", "filePath": "config.xlsx"}
```

## 🏗️ Architecture

### Core Components

```text
ExcelMcp.McpServer/
├── Tools/
│   └── ExcelTools.cs        # 6 resource-based MCP tools  
├── Program.cs               # Official MCP SDK hosting
└── ExcelMcp.McpServer.csproj
```

### Dependencies

```text
ExcelMcp.McpServer
├── ExcelMcp.Core            # Shared Excel automation logic
├── ModelContextProtocol     # Official MCP SDK (v0.4.0-preview.2)  
└── Microsoft.Extensions.*   # Hosting, Logging, DI
```

### Design Patterns

- **Official MCP SDK** - Uses Microsoft's official ModelContextProtocol NuGet package
- **Resource-Based Architecture** - 6 tools instead of 33+ granular operations  
- **Action Pattern** - Each tool supports multiple actions (REST-like design)
- **Attribute-Based Registration** - `[McpServerTool]` and `[McpServerToolType]` attributes
- **JSON Serialization** - Proper `JsonSerializer.Serialize()` for all responses
- **COM Lifecycle Management** - Leverages ExcelMcp.Core's proven Excel automation

## 🔧 System Requirements

| Requirement | Reason |
|-------------|---------|
| **Windows OS** | COM interop for Excel automation |
| **Microsoft Excel** | Direct Excel application control |
| **.NET 10 SDK** | Required for dnx command |
| **ExcelMcp.Core** | Shared Excel automation logic |

## 🔍 Protocol Details

### MCP Protocol Implementation

- **SDK**: Official ModelContextProtocol NuGet package v0.4.0-preview.2
- **Transport**: stdio (stdin/stdout) via `WithStdioServerTransport()`
- **Registration**: Attribute-based tool discovery via `WithToolsFromAssembly()`
- **Hosting**: Microsoft.Extensions.Hosting with console logging to stderr

### Resource-Based Tool Schema

Each tool follows a consistent action-based pattern:

```json
{
  "name": "excel_file",
  "description": "Create, validate, and manage Excel files (.xlsx, .xlsm). Supports actions: create-empty, validate, check-exists.",
  "inputSchema": {
    "type": "object", 
    "properties": {
      "action": {
        "type": "string",
        "description": "Action to perform: create-empty, validate, check-exists"
      },
      "filePath": {
        "type": "string", 
        "description": "Excel file path (.xlsx or .xlsm extension)"
      },
      "macroEnabled": {
        "type": "boolean",
        "description": "Optional: macro-enabled flag for create-empty (default: false)"
      }
    },
    "required": ["action", "filePath"]
  }
}
```

### JSON Response Format

**Success Response:**

```json
{
  "success": true,
  "action": "create-empty", 
  "filePath": "report.xlsx",
  "message": "Excel file created successfully"
}
```

**Error Response:**  

```json
{
  "error": "File not found",
  "action": "validate",
  "filePath": "missing.xlsx"
}
```

## 🚨 Security Considerations

- **File Path Validation** - Prevents path traversal attacks
- **Excel File Validation** - Only processes valid Excel files
- **COM Object Management** - Proper resource cleanup
- **Error Sanitization** - No sensitive information in error messages

### Development Workflow Examples with GitHub Copilot

**Refactoring Power Query M Code:**

```text
User: "This Power Query is slow and hard to read. Can you help me refactor it?"

Copilot: I'll help you refactor that Power Query for better performance and readability:

1. First, let me examine the current M code structure
2. I'll identify performance bottlenecks and complex logic
3. Then refactor with proper step naming and optimization
4. Finally, test the refactored query maintains the same output

[Executes: excel_powerquery view -> analyze code -> excel_powerquery update -> excel_powerquery refresh]

Your Power Query is now optimized with:
- Clearer step names and documentation
- Reduced data loading operations  
- Better error handling patterns
- Improved performance through query folding
```

**VBA Code Development and Debugging:**

```text
User: "I need to add error handling and logging to this VBA module"

Copilot: I'll enhance your VBA code with proper error handling and logging:

1. Export current VBA module for analysis
2. Add comprehensive error handling patterns
3. Implement logging functionality  
4. Update the module with improved code
5. Test the enhanced functionality

[Executes: excel_vba export -> analyze code -> enhance with error handling -> excel_vba update]

Your VBA module now includes:
- Try-catch error handling with proper cleanup
- Structured logging for debugging
- Input validation and sanitization
- Better code organization and comments
```

**Power Query Code Review and Optimization:**

```text
User: "Review this Power Query and suggest performance improvements"

Copilot: I'll analyze your Power Query M code for optimization opportunities:

1. Export and examine the current M code structure
2. Identify query folding issues and inefficient operations
3. Suggest better data source patterns
4. Recommend caching and filtering strategies

[Executes: excel_powerquery view -> code analysis -> optimization suggestions]

Code review findings:
- Move filters closer to source for better query folding
- Replace nested Table.SelectRows with single filter operation  
- Add proper data type conversions early in the pipeline
- Use Table.Buffer strategically for repeated operations
```

## 📚 Documentation

- **[Main ExcelMcp Project](../../../README.md)** - CLI tools overview and installation
- **[Command Reference](../../../docs/COMMANDS.md)** - Complete CLI command documentation
- **[NuGet Trusted Publishing Guide](../../../docs/NUGET_TRUSTED_PUBLISHING.md)** - Setup guide for maintainers

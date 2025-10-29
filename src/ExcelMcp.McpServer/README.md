# ExcelMcp Model Context Protocol (MCP) Server

<!-- mcp-name: io.github.sbroenne/mcp-server-excel -->

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

The MCP server provides **10 focused resource-based tools** optimized for AI coding agents. Each tool handles only Excel-specific operations:

### 1. **`excel_file`** - Excel File Creation 🎯

**Actions**: `create-empty` (1 action)

- Create new Excel workbooks (.xlsx or .xlsm) for automation workflows
- 🎯 **LLM-Optimized**: File validation and existence checks can be done natively by AI agents

### 2. **`excel_powerquery`** - Power Query M Code Management 🧠

**Actions**: `list`, `view`, `import`, `export`, `update`, `delete`, `set-load-to-table`, `set-load-to-data-model`, `set-load-to-both`, `set-connection-only`, `get-load-config` (11 actions)

- Complete Power Query lifecycle for AI-assisted M code development
- Import/export queries for version control and code review
- Configure data loading modes and refresh connections
- 🎯 **LLM-Optimized**: AI can analyze and refactor M code for performance

### 3. **`excel_connection`** - Data Connection Management 🔌

**Actions**: `list`, `view`, `import`, `export`, `update`, `refresh`, `delete`, `loadto`, `properties`, `set-properties`, `test` (11 actions)

- Manage OLEDB, ODBC, Text, Web, and other Excel data connections
- Import/export connection definitions for version control
- Configure connection properties (background query, refresh settings, password handling)
- 🎯 **LLM-Optimized**: AI can manage external data sources and refresh strategies
- 🔒 **Security**: Automatic password sanitization in all outputs

### 4. **`excel_datamodel`** - Data Model & DAX Management 📈

**Actions**: `list-tables`, `list-measures`, `view-measure`, `export-measure`, `list-relationships`, `refresh`, `delete-measure`, `delete-relationship`, `list-columns`, `view-table`, `get-model-info`, `create-measure`, `update-measure`, `create-relationship`, `update-relationship`, `create-column`, `view-column`, `update-column`, `delete-column`, `validate-dax` (20 actions)

- Excel Data Model (Power Pivot) operations for enterprise analytics
- **Phase 2 CRUD:** Create/update DAX measures with format types (Currency, Percentage, Decimal, General)
- **Phase 2 CRUD:** Create/update table relationships with active/inactive toggles
- **Phase 2 CRUD:** Create/update/delete DAX calculated columns with data types (String, Integer, Double, Boolean, DateTime)
- Discovery: List tables, columns, measures, relationships; view table details; get model statistics
- DAX measure inspection, export to .dax files for version control, and deletion
- Table relationship management and analysis
- Data Model refresh and structure exploration
- DAX formula validation without creating objects
- 🎯 **LLM-Optimized**: AI can create/manage DAX formulas, calculated columns, and optimize Data Model structure

### 5. **`excel_worksheet`** - Worksheet Lifecycle Management 📊

**Actions**: `list`, `create`, `rename`, `copy`, `delete` (5 actions)  

- Worksheet lifecycle management (creation, renaming, copying, deletion)
- List all worksheets with visibility and index information
- 📝 **Note**: Data operations (read/write/clear) moved to `excel_range` in Phase 1
- 🎯 **LLM-Optimized**: AI manages worksheet structure while using `excel_range` for data

### 6. **`excel_range`** - Unified Range Operations 🎯

**Actions**: `get-values`, `set-values`, `get-formulas`, `set-formulas`, `clear-all`, `clear-contents`, `clear-formats`, and 23+ more (30+ actions)

- **Unified range API**: Single cell = 1x1 range (e.g., "A1" returns `[[value]]`)
- **Phase 1 Consolidation:** Replaced `excel_cell` tool, absorbed worksheet data operations
- Value/formula operations with 2D arrays (JSON format)
- Clear, copy, insert/delete cells/rows/columns
- Find/replace, sorting, hyperlink management
- Discovery: UsedRange, CurrentRegion, RangeInfo
- Named ranges: Transparent resolution (accepts both "A1:D10" and "SalesData")
- 🎯 **LLM-Optimized**: AI works with ranges naturally without cell vs range confusion

### 7. **`excel_parameter`** - Named Ranges as Configuration ⚙️

**Actions**: `list`, `get`, `set`, `create`, `delete`, `update` (6 actions)

- Excel configuration management through named ranges for dynamic AI-controlled parameters
- Parameter-driven workbook automation and templating
- Update named range references to point to different cells
- 🎯 **LLM-Optimized**: AI can dynamically configure Excel behavior via parameters

### 8. **`excel_table`** - Excel Table (ListObject) Management 📊

**Actions**: `list`, `create`, `info`, `rename`, `delete`, `resize`, `set-style`, `toggle-totals`, `set-column-total`, `append`, `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `add-to-datamodel` (22 actions)

- **Phase 2 Advanced Features:** Complete Excel Table automation
- Table lifecycle: Create, rename, delete, resize with auto-expansion
- Structure: Add/remove/rename columns, set visual styles
- Totals: Toggle totals row, set column total functions (SUM, AVG, COUNT, etc.)
- Filtering: Apply criteria filters, value filters, clear filters, get filter state
- Sorting: Single column or multi-column sorts (up to 3 levels)
- Structured References: Get Excel formulas for table regions (All, Data, Headers, Totals, ThisRow)
- Data Model Integration: Add tables to Power Pivot for analytics
- **Standalone Features**: AutoFilter dropdowns, structured references ([@Column]), dynamic expansion, formatted bands
- **Power Query Integration**: Tables can optionally be referenced in Power Query queries
- 🎯 **LLM-Optimized**: AI can build complex table structures with filters, sorts, and formulas

### 9. **`excel_vba`** - VBA Macro Management & Execution 📜

**Actions**: `list`, `view`, `export`, `import`, `update`, `run`, `delete` (7 actions) ⚠️ *(.xlsm files only)*

- Complete VBA lifecycle for AI-assisted macro development and automation
- View module code without exporting to file
- Script import/export for version control and code review
- Execute macros with parameters directly from AI
- 🎯 **LLM-Optimized**: AI can enhance VBA with error handling, logging, and best practices

### 10. **`excel_version`** - Version Checking ⚡

**Actions**: `check` (1 action)

- Check for updates on NuGet.org
- Automatic startup check displays warning if outdated version detected
- Provides update instructions and workflow guidance
- 🎯 **LLM-Optimized**: AI can notify users about available updates and guide upgrade process

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

### Connection Management

```text
User: "Show me all data connections in my workbook"
AI Assistant uses: excel_connection(action="list", excelPath="data-analysis.xlsx")
Result: {"success": true, "connections": [{"name": "SalesDB", "type": "OLEDB", "isPowerQuery": false}]}

User: "Refresh the SalesDB connection to get latest data"
AI Assistant uses: excel_connection(action="refresh", excelPath="data-analysis.xlsx", connectionName="SalesDB")
Result: {"success": true, "message": "Connection 'SalesDB' refreshed successfully"}

User: "Export the connection definition for version control"
AI Assistant uses: excel_connection(action="export", excelPath="data-analysis.xlsx", connectionName="SalesDB", targetPath="salesdb-connection.json")
Result: {"success": true, "message": "Connection exported to salesdb-connection.json"}
```

### Data Model Management

```text
User: "Show me all tables in the Data Model"
AI Assistant uses: excel_datamodel(action="list-tables", excelPath="sales-analysis.xlsx")
Result: {"success": true, "tables": [{"name": "Sales", "recordCount": 15420}, {"name": "Customers", "recordCount": 350}]}

User: "List all DAX measures in my workbook"
AI Assistant uses: excel_datamodel(action="list-measures", excelPath="sales-analysis.xlsx")
Result: {"success": true, "measures": [{"name": "Total Sales", "formula": "SUM(Sales[Amount])", "table": "Sales"}]}

User: "Create a new DAX measure for total revenue with currency formatting"
AI Assistant uses: excel_datamodel(action="create-measure", excelPath="sales.xlsx", tableName="Sales", measureName="TotalRevenue", formula="SUM(Sales[Amount])", formatType="Currency", description="Total sales revenue")
Result: {"success": true, "message": "Measure 'TotalRevenue' created successfully"}

User: "Update the Total Sales formula to use CALCULATE"
AI Assistant uses: excel_datamodel(action="update-measure", excelPath="sales.xlsx", measureName="Total Sales", formula="CALCULATE(SUM(Sales[Amount]))", description="Updated with CALCULATE for better performance")
Result: {"success": true, "message": "Measure 'Total Sales' updated successfully"}

User: "Create a relationship between Sales and Customers tables"
AI Assistant uses: excel_datamodel(action="create-relationship", excelPath="sales.xlsx", fromTable="Sales", fromColumn="CustomerID", toTable="Customers", toColumn="ID", active=true)
Result: {"success": true, "message": "Relationship created successfully"}

User: "Export the 'Total Sales' measure to a file for version control"
AI Assistant uses: excel_datamodel(action="export-measure", excelPath="sales-analysis.xlsx", measureName="Total Sales", outputPath="measures/total-sales.dax")
Result: {"success": true, "message": "Measure exported successfully"}

User: "Show me all relationships between tables"
AI Assistant uses: excel_datamodel(action="list-relationships", excelPath="sales-analysis.xlsx")
Result: {"success": true, "relationships": [{"fromTable": "Sales", "fromColumn": "CustomerID", "toTable": "Customers", "toColumn": "ID", "isActive": true}]}

User: "Delete the old 'Previous Year Sales' measure"
AI Assistant uses: excel_datamodel(action="delete-measure", excelPath="sales-analysis.xlsx", measureName="Previous Year Sales")
Result: {"success": true, "message": "Measure 'Previous Year Sales' deleted successfully"}

User: "Refresh the Data Model to get latest data"
AI Assistant uses: excel_datamodel(action="refresh", excelPath="sales-analysis.xlsx")
Result: {"success": true, "message": "Data Model refreshed successfully"}

User: "Create a calculated column for total cost in the Sales table"
AI Assistant uses: excel_datamodel(action="create-column", excelPath="sales.xlsx", tableName="Sales", columnName="TotalCost", daxFormula="[Price] * [Quantity]", dataType="Double", description="Total cost per item")
Result: {"success": true, "message": "Calculated column 'TotalCost' created successfully"}

User: "List all calculated columns in my Data Model"
AI Assistant uses: excel_datamodel(action="list-columns", excelPath="sales.xlsx")
Result: {"success": true, "calculatedColumns": [{"name": "TotalCost", "table": "Sales", "dataType": "Double", "formulaPreview": "[Price] * [Quantity]"}]}

User: "Show me the details for the Profit calculated column"
AI Assistant uses: excel_datamodel(action="view-column", excelPath="sales.xlsx", tableName="Sales", columnName="Profit")
Result: {"success": true, "columnName": "Profit", "tableName": "Sales", "daxFormula": "[Revenue] - [Cost]", "dataType": "Double"}

User: "Update the TotalCost column to include tax"
AI Assistant uses: excel_datamodel(action="update-column", excelPath="sales.xlsx", tableName="Sales", columnName="TotalCost", daxFormula="[Price] * [Quantity] * 1.1", description="Total cost with 10% tax")
Result: {"success": true, "message": "Calculated column 'TotalCost' updated successfully"}

User: "Validate this DAX formula before creating a column"
AI Assistant uses: excel_datamodel(action="validate-dax", excelPath="sales.xlsx", daxFormula="[Revenue] - [Cost]")
Result: {"success": true, "isValid": true, "daxFormula": "[Revenue] - [Cost]"}
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

### Range & Cell Operations

```text  
User: "What's the value in cell B5 of the Summary sheet?"
AI Assistant uses: excel_range(action="get-values", excelPath="report.xlsx", sheetName="Summary", rangeAddress="B5")
Result: {"success": true, "values": [[42]], "sheetName": "Summary", "rangeAddress": "B5"}

User: "Set cell A1 to contain the formula =SUM(B1:B10)"  
AI Assistant uses: excel_range(action="set-formulas", excelPath="report.xlsx", sheetName="Sheet1", rangeAddress="A1", formulas=[["=SUM(B1:B10)"]])
Result: {"success": true, "action": "set-formulas", "sheetName": "Sheet1", "rangeAddress": "A1"}

User: "Read the sales data from range A1:D100"
AI Assistant uses: excel_range(action="get-values", excelPath="sales.xlsx", sheetName="Data", rangeAddress="A1:D100")
Result: {"success": true, "values": [[...2D array...]], "rowCount": 100, "columnCount": 4}
```

### Table Operations

```text
User: "List all Excel Tables in my workbook"
AI Assistant uses: excel_table(action="list", filePath="sales.xlsx")
Result: {"success": true, "tables": [{"name": "SalesTable", "sheetName": "Data", "recordCount": 150}]}

User: "Apply a filter to show only sales over $100"
AI Assistant uses: excel_table(action="apply-filter", filePath="sales.xlsx", tableName="SalesTable", columnName="Amount", criteria=">100")
Result: {"success": true, "message": "Filter applied successfully"}

User: "Sort the table by Amount descending, then by Date ascending"
AI Assistant uses: excel_table(action="sort-multi", filePath="sales.xlsx", tableName="SalesTable", sortColumns=[{"columnName": "Amount", "ascending": false}, {"columnName": "Date", "ascending": true}])
Result: {"success": true, "message": "Table sorted successfully"}
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

## 📚 Reference Prompts

The MCP server provides **1 reference prompt** to help AI assistants understand Excel connection types and COM API limitations:

### Connection Type Reference
- **`excel_connection_reference`** - Quick reference for Excel's 9 connection types, which ones work via COM API, and critical limitations

**Why This Prompt Exists:**
- Excel connection types and COM API quirks are **niche domain knowledge** not in standard LLM training
- Prevents incorrect answers (e.g., telling users to create OLEDB connections via COM API, which fails)
- Explains known issues like Type 3/4 confusion when users encounter them

**Usage:** AI assistants can invoke this prompt when users ask about connection types, connection failures, or data source management.

**Example Invocation:**
```text
User: "Why won't my SQL Server connection work?"
AI: Get prompt excel_connection_reference
AI: "Excel COM API cannot create OLEDB connections via Connections.Add(). 
     Create the connection in Excel UI (Data → Get Data → From Database), 
     then use excel_connection to manage it."
```

## 🏗️ Architecture

### Core Components

```text
ExcelMcp.McpServer/
├── Tools/
│   └── ExcelTools.cs        # 9 resource-based MCP tools  
├── Prompts/
│   └── ExcelConnectionPrompts.cs    # Connection type reference
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
- **Resource-Based Architecture** - 9 tools instead of 56+ granular operations  
- **Action Pattern** - Each tool supports multiple actions (REST-like design)
- **Attribute-Based Registration** - `[McpServerTool]`, `[McpServerPrompt]` attributes for auto-discovery
- **Reference Prompt** - 1 prompt provides Excel connection type reference for AI assistants
- **JSON Serialization** - Proper `JsonSerializer.Serialize()` for all responses
- **COM Lifecycle Management** - Leverages ExcelMcp.Core's proven Excel automation

## 🔧 System Requirements

| Requirement | Reason |
|-------------|---------|
| **Windows OS** | COM interop for Excel automation |
| **Microsoft Excel** | Direct Excel application control |
| **.NET SDK** | Required for `dnx` command (any recent version) |
| **.NET 8.0 Runtime** | MCP server targets .NET 8.0 |
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

**Version Check and Update Workflow:**

```text
User: "Check if I'm running the latest version of ExcelMcp"

Copilot: Let me check for updates:

[Executes: excel_version check]

Result: A newer version (1.2.0) is available. You are running version 1.0.0.

Update instructions:
The dnx command automatically downloads the latest version from NuGet.
Simply restart VS Code to update - the MCP server will use the new version automatically.

Would you like me to verify the update after you restart?
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

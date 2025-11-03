# ExcelMcp.Core

**Core library for Excel automation operations via COM interop by Sbroenne.**

## Overview

`Sbroenne.ExcelMcp.Core` provides high-level Excel automation commands for Power Query, VBA, Data Model (Power Pivot), worksheets, ranges, tables, parameters, and connections. It's the business logic layer shared by ExcelMcp.McpServer and ExcelMcp.CLI.

## Features

### Power Query (16 operations)
- List, view, create, update, delete Power Queries
- Import/export M code for version control
- Manage load destinations (worksheet, data model, connection-only)
- Configure privacy levels and refresh settings

### Data Model / Power Pivot (15 operations)
- Create and manage DAX measures
- Define relationships between tables
- Export measures to .dax files
- Discover model structure (tables, columns, measures)

### VBA Macros (7 operations)
- List, export, import VBA modules
- Run VBA procedures
- Integrate with version control systems

### Excel Tables (26 operations)
- Complete table lifecycle management
- Filtering, sorting, styling
- Column management and structured references
- Data appending and manipulation

### Ranges & Data (45 operations)
- Get/set values and formulas
- Copy, paste, find, replace operations
- Formatting and data validation
- Cell protection and conditional formatting

### Worksheets (13 operations)
- Create, rename, copy, delete worksheets
- Tab colors and visibility controls
- Worksheet navigation

### Connections (11 operations)
- Manage OLEDB, ODBC, Text, Web connections
- Import/export connection files (.odc)
- Refresh and test connections

### Named Ranges (7 operations)
- Create and manage named range parameters
- Bulk operations on parameters
- Dynamic range definitions

## Installation

```bash
dotnet add package Sbroenne.ExcelMcp.Core
```

## Usage Example

```csharp
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop;

// Create command instances
var powerQueryCommands = new PowerQueryCommands();
var sheetCommands = new SheetCommands();

// Use batch operations for efficiency
await using var batch = await ExcelSession.BeginBatchAsync("workbook.xlsx");

// List Power Queries
var queries = await powerQueryCommands.ListAsync(batch);
foreach (var query in queries.Items)
{
    Console.WriteLine($"Query: {query.Name}");
}

// Create a new worksheet
var result = await sheetCommands.CreateAsync(batch, "NewSheet");
if (result.Success)
{
    Console.WriteLine("Sheet created successfully");
}

await batch.SaveAsync();
```

## Architecture

```
ExcelMcp.Core (Business Logic)
    └─ ExcelMcp.ComInterop (COM Interop Layer)
```

## Requirements

- Windows OS
- .NET 8.0 or later
- Microsoft Excel 2016+ installed
- Sbroenne.ExcelMcp.ComInterop (included as dependency)

## Platform Support

- ✅ Windows x64
- ✅ Windows ARM64
- ❌ Linux (Excel COM not available)
- ❌ macOS (Excel COM not available)

## Related Packages

- **Sbroenne.ExcelMcp.ComInterop** - Low-level COM interop utilities (dependency)
- **Sbroenne.ExcelMcp.McpServer** - MCP server for AI assistant integration
- **Sbroenne.ExcelMcp.CLI** - Command-line tool for Excel automation

## Documentation

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Commands Reference](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/COMMANDS.md)

## License

MIT License - see LICENSE file for details.

## Repository

https://github.com/sbroenne/mcp-server-excel

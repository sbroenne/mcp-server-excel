# ExcelMcp.ComInterop

**Low-level COM interop utilities for Excel automation by Sbroenne.**

## Overview

`Sbroenne.ExcelMcp.ComInterop` is a reusable library providing robust COM object lifecycle management and OLE message filtering for Excel automation. It's the foundation layer for ExcelMcp projects but can be used independently in any .NET application that automates Excel via COM interop.

## Features

- **STA Threading Management** - Ensures proper single-threaded apartment model for Excel COM objects
- **COM Object Lifecycle** - Automatic COM object cleanup and garbage collection
- **OLE Message Filtering** - Handles busy/rejected COM calls with retry logic using Polly
- **Excel Session Management** - Manages Excel.Application lifecycle safely
- **Batch Operations** - Efficient handling of multiple Excel operations in a single session

## Installation

```bash
dotnet add package Sbroenne.ExcelMcp.ComInterop
```

## Usage Example

```csharp
using Sbroenne.ExcelMcp.ComInterop;

// Use ExcelSession for safe Excel automation
await using var session = await ExcelSession.BeginAsync("path/to/workbook.xlsx");
await using var batch = await session.BeginBatchAsync();

await batch.ExecuteAsync<int>(async (ctx, ct) => 
{
    // Access Excel workbook through ctx.Book
    dynamic worksheets = ctx.Book.Worksheets;
    dynamic sheet = worksheets.Item[1];
    
    // Perform Excel operations
    sheet.Name = "UpdatedSheet";
    
    return 0;
});

await batch.SaveAsync();
```

## Key Classes

- **ExcelSession** - Manages Excel.Application lifecycle and workbook operations
- **ExcelBatch** - Groups multiple operations for efficient execution
- **ComUtilities** - Helper methods for COM object cleanup and safe property access
- **OleMessageFilter** - Implements retry logic for busy Excel instances

## Requirements

- Windows OS
- .NET 8.0 or later
- Microsoft Excel 2016+ installed

## Platform Support

- ✅ Windows x64
- ✅ Windows ARM64
- ❌ Linux (Excel COM not available)
- ❌ macOS (Excel COM not available)

## Related Packages

- **Sbroenne.ExcelMcp.Core** - High-level Excel automation commands
- **Sbroenne.ExcelMcp.McpServer** - MCP server for AI assistant integration
- **Sbroenne.ExcelMcp.CLI** - Command-line tool for Excel automation

## License

MIT License - see LICENSE file for details.

## Repository

https://github.com/sbroenne/mcp-server-excel

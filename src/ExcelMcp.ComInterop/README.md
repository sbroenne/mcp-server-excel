# ExcelMcp.ComInterop

**Low-level COM interop utilities for Excel automation by Sbroenne.**

## Overview

This library provides Excel-specific COM object lifecycle management and OLE message filtering. It's the foundation layer for ExcelMcp projects, handling STA threading, session management, and batch operations specifically for Excel COM automation.

**Note:** Despite the generic name "ComInterop", this library is Excel-specific and not intended for Word/PowerPoint/other Office applications.

## Features

- **STA Threading Management** - Ensures proper single-threaded apartment model for Excel COM objects
- **COM Object Lifecycle** - Automatic COM object cleanup and garbage collection
- **OLE Message Filtering** - Handles busy/rejected COM calls with retry logic using Polly
- **Excel Session Management** - Manages Excel.Application lifecycle safely
- **Batch Operations** - Efficient handling of multiple Excel operations in a single session

## Usage Example

```csharp
using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

using var batch = ExcelSession.BeginBatch(@"C:\Data\workbook.xlsx");

batch.Execute((ctx, ct) =>
{
    Excel.Sheets? worksheets = null;
    Excel.Worksheet? sheet = null;

    try
    {
        worksheets = ctx.Book.Worksheets;
        sheet = (Excel.Worksheet)worksheets[1];
        sheet.Name = "UpdatedSheet";
    }
    finally
    {
        ComUtilities.Release(ref sheet);
        ComUtilities.Release(ref worksheets);
    }
});

batch.Save();
```

## Key Classes

- **ExcelSession** - Manages Excel.Application lifecycle and workbook operations
- **ExcelBatch** - Groups multiple operations for efficient execution
- **ComUtilities** - Helper methods for COM object cleanup and safe property access
- **OleMessageFilter** - Implements retry logic for busy Excel instances

## Requirements

- Windows OS
- .NET 10.0 or later
- Microsoft Excel 2016+ installed

## Platform Support

- ✅ Windows x64
- ✅ Windows ARM64
- ❌ Linux (Excel COM not available)
- ❌ macOS (Excel COM not available)

# ExcelMcp CLI Examples

This directory contains example scripts demonstrating ExcelMcp CLI features.

## Session Mode Demo

The session mode demo shows how to use sessions for high-performance multi-operation workflows.

### Requirements

- Windows with Excel installed
- ExcelMcp installed (`dotnet tool install --global Sbroenne.ExcelMcp.McpServer`)

### Running the Demo

**Linux/macOS/WSL:**
```powershell
./session-demo.sh
```

**Windows PowerShell:**
```powershell
.\session-demo.ps1
```

### What the Demo Does

1. Creates a test workbook (`test-session.xlsx`)
2. Opens a session and captures the session ID
3. Performs multiple operations using the same Excel instance:
   - Creates 3 worksheets (Sales, Customers, Products)
   - Lists worksheets
   - Lists Power Queries
4. Lists active sessions
5. Closes the session with `--save` (saves all changes)
6. Verifies changes were saved

### Expected Performance

Session mode is **75-90% faster** than running individual commands because:
- Only one Excel instance is opened
- No file open/close overhead between operations
- All changes committed atomically

### Cleanup

```powershell
rm test-session.xlsx
```

Or in PowerShell:
```powershell
Remove-Item test-session.xlsx
```

## Use Cases

Session mode is ideal for:
- **RPA workflows** - Automated report generation
- **Data pipelines** - ETL operations with multiple steps
- **Testing** - Setting up test data across multiple sheets
- **Bulk operations** - Making many changes to a workbook

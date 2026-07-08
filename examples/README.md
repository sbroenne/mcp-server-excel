# ExcelMcp CLI Examples

This directory contains example scripts demonstrating ExcelMcp CLI features.

## Session Mode Demo

The session mode demo shows how to use sessions for high-performance multi-operation workflows.

### Requirements

- Windows with Excel installed
- ExcelMcp installed (standalone exe from [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) or via `dotnet tool install --global Sbroenne.ExcelMcp.McpServer`)

### Running the Demo

Run these commands from PowerShell on Windows:

```powershell
# 1. Create a new workbook and open a session (captures the session ID)
$session = (excelcli session create test-session.xlsx | Select-String "Session ID:").ToString().Split()[-1]

# 2. Perform multiple operations against the same Excel instance
excelcli sheet create --session $session --sheet Sales
excelcli sheet create --session $session --sheet Customers
excelcli sheet create --session $session --sheet Products
excelcli sheet list --session $session
excelcli powerquery list --session $session

# 3. List active sessions
excelcli session list

# 4. Save all changes and close the session
excelcli session close --session $session --save
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
Remove-Item test-session.xlsx
```

## Use Cases

Session mode is ideal for:
- **RPA workflows** - Automated report generation
- **Data pipelines** - ETL operations with multiple steps
- **Testing** - Setting up test data across multiple sheets
- **Bulk operations** - Making many changes to a workbook

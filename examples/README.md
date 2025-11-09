# ExcelMcp.CLI Examples

This directory contains example scripts demonstrating ExcelMcp.CLI features.

## Batch Mode Demo

The batch mode demo shows how to use batch sessions for high-performance multi-operation workflows.

### Requirements

- Windows with Excel installed
- ExcelMcp.CLI installed (`dotnet tool install --global Sbroenne.ExcelMcp.CLI`)

### Running the Demo

**Linux/macOS/WSL:**
```bash
./batch-mode-demo.sh
```

**Windows PowerShell:**
```powershell
.\batch-mode-demo.ps1
```

### What the Demo Does

1. Creates a test workbook (`test-batch.xlsx`)
2. Starts a batch session and captures the batch ID
3. Performs multiple operations using the same Excel instance:
   - Creates 3 worksheets (Sales, Customers, Products)
   - Lists worksheets
   - Lists Power Queries
4. Lists active batch sessions
5. Commits the batch (saves all changes)
6. Verifies changes were saved

### Expected Performance

Batch mode is **75-90% faster** than running individual commands because:
- Only one Excel instance is opened
- No file open/close overhead between operations
- All changes committed atomically

### Cleanup

```bash
rm test-batch.xlsx
```

Or in PowerShell:
```powershell
Remove-Item test-batch.xlsx
```

## Use Cases

Batch mode is ideal for:
- **RPA workflows** - Automated report generation
- **Data pipelines** - ETL operations with multiple steps
- **Testing** - Setting up test data across multiple sheets
- **Bulk operations** - Making many changes to a workbook

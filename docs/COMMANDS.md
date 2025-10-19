# Command Reference

Complete reference for all ExcelMcp.CLI commands.

## File Commands

Create and manage Excel workbooks.

**create-empty** - Create empty Excel workbook

```powershell
ExcelMcp.CLI create-empty <file.xlsx|file.xlsm>
```

Essential for coding agents to programmatically create Excel workbooks. Use `.xlsm` extension for macro-enabled workbooks that support VBA.

## Power Query Commands (`pq-*`)

Manage Power Query queries in Excel workbooks.

**pq-list** - List all Power Queries

```powershell
ExcelMcp.CLI pq-list <file.xlsx>
```

**pq-view** - View Power Query M code

```powershell
ExcelMcp.CLI pq-view <file.xlsx> <query-name>
```

**pq-import** - Create or import query from file

```powershell
ExcelMcp.CLI pq-import <file.xlsx> <query-name> <source.pq>
```

**pq-export** - Export query to file

```powershell
ExcelMcp.CLI pq-export <file.xlsx> <query-name> <output.pq>
```

**pq-update** - Update existing query from file

```powershell
ExcelMcp.CLI pq-update <file.xlsx> <query-name> <code.pq>
```

**pq-refresh** - Refresh query data

```powershell
ExcelMcp.CLI pq-refresh <file.xlsx> <query-name>
```

**pq-loadto** - Load connection-only query to worksheet

```powershell
ExcelMcp.CLI pq-loadto <file.xlsx> <query-name> <sheet-name>
```

**pq-delete** - Delete Power Query

```powershell
ExcelMcp.CLI pq-delete <file.xlsx> <query-name>
```

## Sheet Commands (`sheet-*`)

Manage worksheets in Excel workbooks.

**sheet-list** - List all worksheets

```powershell
ExcelMcp.CLI sheet-list <file.xlsx>
```

**sheet-read** - Read data from worksheet

```powershell
ExcelMcp.CLI sheet-read <file.xlsx> <sheet-name> [range]

# Examples
ExcelMcp.CLI sheet-read "Plan.xlsx" "Data"           # Read entire used range
ExcelMcp.CLI sheet-read "Plan.xlsx" "Data" "A1:C10"  # Read specific range
```

**sheet-write** - Write CSV data to worksheet

```powershell
ExcelMcp.CLI sheet-write <file.xlsx> <sheet-name> <data.csv>
```

**sheet-create** - Create new worksheet

```powershell
ExcelMcp.CLI sheet-create <file.xlsx> <sheet-name>
```

**sheet-copy** - Copy worksheet

```powershell
ExcelMcp.CLI sheet-copy <file.xlsx> <source-sheet> <new-sheet>
```

**sheet-rename** - Rename worksheet

```powershell
ExcelMcp.CLI sheet-rename <file.xlsx> <old-name> <new-name>
```

**sheet-delete** - Delete worksheet

```powershell
ExcelMcp.CLI sheet-delete <file.xlsx> <sheet-name>
```

**sheet-clear** - Clear worksheet data

```powershell
ExcelMcp.CLI sheet-clear <file.xlsx> <sheet-name> [range]
```

**sheet-append** - Append CSV data to worksheet

```powershell
ExcelMcp.CLI sheet-append <file.xlsx> <sheet-name> <data.csv>
```

## Parameter Commands (`param-*`)

Manage named ranges and parameters.

**param-list** - List all named ranges

```powershell
ExcelMcp.CLI param-list <file.xlsx>
```

**param-get** - Get named range value

```powershell
ExcelMcp.CLI param-get <file.xlsx> <param-name>
```

**param-set** - Set named range value

```powershell
ExcelMcp.CLI param-set <file.xlsx> <param-name> <value>
```

**param-create** - Create named range

```powershell
ExcelMcp.CLI param-create <file.xlsx> <param-name> <reference>
```

**param-delete** - Delete named range

```powershell
ExcelMcp.CLI param-delete <file.xlsx> <param-name>
```

## Cell Commands (`cell-*`)

Manage individual cells.

**cell-get-value** - Get cell value

```powershell
ExcelMcp.CLI cell-get-value <file.xlsx> <sheet-name> <cell>
```

**cell-set-value** - Set cell value

```powershell
ExcelMcp.CLI cell-set-value <file.xlsx> <sheet-name> <cell> <value>
```

**cell-get-formula** - Get cell formula

```powershell
ExcelMcp.CLI cell-get-formula <file.xlsx> <sheet-name> <cell>
```

**cell-set-formula** - Set cell formula

```powershell
ExcelMcp.CLI cell-set-formula <file.xlsx> <sheet-name> <cell> <formula>
```

## VBA Script Commands (`script-*`)

**⚠️ VBA commands require macro-enabled (.xlsm) files!**

Manage VBA scripts and macros in macro-enabled Excel workbooks.

**script-list** - List all VBA modules and procedures

```powershell
ExcelMcp.CLI script-list <file.xlsm>
```

**script-export** - Export VBA module to file

```powershell
ExcelMcp.CLI script-export <file.xlsm> <module-name> <output.vba>
```

**script-import** - Import VBA module from file

```powershell
ExcelMcp.CLI script-import <file.xlsm> <module-name> <source.vba>
```

**script-update** - Update existing VBA module

```powershell
ExcelMcp.CLI script-update <file.xlsm> <module-name> <source.vba>
```

**script-run** - Execute VBA macro with parameters

```powershell
ExcelMcp.CLI script-run <file.xlsm> <macro-name> [param1] [param2] ...

# Examples
ExcelMcp.CLI script-run "Report.xlsm" "ProcessData"
ExcelMcp.CLI script-run "Analysis.xlsm" "CalculateTotal" "Sheet1" "A1:C10"
```

**script-delete** - Remove VBA module

```powershell
ExcelMcp.CLI script-delete <file.xlsm> <module-name>
```

## Setup Commands

Configure VBA trust settings for automation.

**setup-vba-trust** - Enable VBA project access

```powershell
ExcelMcp.CLI setup-vba-trust
```

**check-vba-trust** - Check VBA trust configuration

```powershell
ExcelMcp.CLI check-vba-trust
```

## File Format Support

- **`.xlsx`** - Standard Excel workbooks (Power Query, worksheets, parameters)
- **`.xlsm`** - Macro-enabled workbooks (includes VBA script support)

Use `create-empty` with `.xlsm` extension to create macro-enabled workbooks:

```powershell
ExcelMcp.CLI create-empty "macros.xlsm"  # Creates macro-enabled workbook
ExcelMcp.CLI create-empty "data.xlsx"    # Creates standard workbook
```

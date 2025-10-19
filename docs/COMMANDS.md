# Command Reference

Complete reference for all ExcelMcp commands.

## File Commands

Create and manage Excel workbooks.

**create-empty** - Create empty Excel workbook

```powershell
ExcelMcp create-empty <file.xlsx|file.xlsm>
```

Essential for coding agents to programmatically create Excel workbooks. Use `.xlsm` extension for macro-enabled workbooks that support VBA.

## Power Query Commands (`pq-*`)

Manage Power Query queries in Excel workbooks.

**pq-list** - List all Power Queries

```powershell
ExcelMcp pq-list <file.xlsx>
```

**pq-view** - View Power Query M code

```powershell
ExcelMcp pq-view <file.xlsx> <query-name>
```

**pq-import** - Create or import query from file

```powershell
ExcelMcp pq-import <file.xlsx> <query-name> <source.pq>
```

**pq-export** - Export query to file

```powershell
ExcelMcp pq-export <file.xlsx> <query-name> <output.pq>
```

**pq-update** - Update existing query from file

```powershell
ExcelMcp pq-update <file.xlsx> <query-name> <code.pq>
```

**pq-refresh** - Refresh query data

```powershell
ExcelMcp pq-refresh <file.xlsx> <query-name>
```

**pq-loadto** - Load connection-only query to worksheet

```powershell
ExcelMcp pq-loadto <file.xlsx> <query-name> <sheet-name>
```

**pq-delete** - Delete Power Query

```powershell
ExcelMcp pq-delete <file.xlsx> <query-name>
```

## Sheet Commands (`sheet-*`)

Manage worksheets in Excel workbooks.

**sheet-list** - List all worksheets

```powershell
ExcelMcp sheet-list <file.xlsx>
```

**sheet-read** - Read data from worksheet

```powershell
ExcelMcp sheet-read <file.xlsx> <sheet-name> [range]

# Examples
ExcelMcp sheet-read "Plan.xlsx" "Data"           # Read entire used range
ExcelMcp sheet-read "Plan.xlsx" "Data" "A1:C10"  # Read specific range
```

**sheet-write** - Write CSV data to worksheet

```powershell
ExcelMcp sheet-write <file.xlsx> <sheet-name> <data.csv>
```

**sheet-create** - Create new worksheet

```powershell
ExcelMcp sheet-create <file.xlsx> <sheet-name>
```

**sheet-copy** - Copy worksheet

```powershell
ExcelMcp sheet-copy <file.xlsx> <source-sheet> <new-sheet>
```

**sheet-rename** - Rename worksheet

```powershell
ExcelMcp sheet-rename <file.xlsx> <old-name> <new-name>
```

**sheet-delete** - Delete worksheet

```powershell
ExcelMcp sheet-delete <file.xlsx> <sheet-name>
```

**sheet-clear** - Clear worksheet data

```powershell
ExcelMcp sheet-clear <file.xlsx> <sheet-name> [range]
```

**sheet-append** - Append CSV data to worksheet

```powershell
ExcelMcp sheet-append <file.xlsx> <sheet-name> <data.csv>
```

## Parameter Commands (`param-*`)

Manage named ranges and parameters.

**param-list** - List all named ranges

```powershell
ExcelMcp param-list <file.xlsx>
```

**param-get** - Get named range value

```powershell
ExcelMcp param-get <file.xlsx> <param-name>
```

**param-set** - Set named range value

```powershell
ExcelMcp param-set <file.xlsx> <param-name> <value>
```

**param-create** - Create named range

```powershell
ExcelMcp param-create <file.xlsx> <param-name> <reference>
```

**param-delete** - Delete named range

```powershell
ExcelMcp param-delete <file.xlsx> <param-name>
```

## Cell Commands (`cell-*`)

Manage individual cells.

**cell-get-value** - Get cell value

```powershell
ExcelMcp cell-get-value <file.xlsx> <sheet-name> <cell>
```

**cell-set-value** - Set cell value

```powershell
ExcelMcp cell-set-value <file.xlsx> <sheet-name> <cell> <value>
```

**cell-get-formula** - Get cell formula

```powershell
ExcelMcp cell-get-formula <file.xlsx> <sheet-name> <cell>
```

**cell-set-formula** - Set cell formula

```powershell
ExcelMcp cell-set-formula <file.xlsx> <sheet-name> <cell> <formula>
```

## VBA Script Commands (`script-*`)

**⚠️ VBA commands require macro-enabled (.xlsm) files!**

Manage VBA scripts and macros in macro-enabled Excel workbooks.

**script-list** - List all VBA modules and procedures

```powershell
ExcelMcp script-list <file.xlsm>
```

**script-export** - Export VBA module to file

```powershell
ExcelMcp script-export <file.xlsm> <module-name> <output.vba>
```

**script-import** - Import VBA module from file

```powershell
ExcelMcp script-import <file.xlsm> <module-name> <source.vba>
```

**script-update** - Update existing VBA module

```powershell
ExcelMcp script-update <file.xlsm> <module-name> <source.vba>
```

**script-run** - Execute VBA macro with parameters

```powershell
ExcelMcp script-run <file.xlsm> <macro-name> [param1] [param2] ...

# Examples
ExcelMcp script-run "Report.xlsm" "ProcessData"
ExcelMcp script-run "Analysis.xlsm" "CalculateTotal" "Sheet1" "A1:C10"
```

**script-delete** - Remove VBA module

```powershell
ExcelMcp script-delete <file.xlsm> <module-name>
```

## Setup Commands

Configure VBA trust settings for automation.

**setup-vba-trust** - Enable VBA project access

```powershell
ExcelMcp setup-vba-trust
```

**check-vba-trust** - Check VBA trust configuration

```powershell
ExcelMcp check-vba-trust
```

## File Format Support

- **`.xlsx`** - Standard Excel workbooks (Power Query, worksheets, parameters)
- **`.xlsm`** - Macro-enabled workbooks (includes VBA script support)

Use `create-empty` with `.xlsm` extension to create macro-enabled workbooks:

```powershell
ExcelMcp create-empty "macros.xlsm"  # Creates macro-enabled workbook
ExcelMcp create-empty "data.xlsx"    # Creates standard workbook
```
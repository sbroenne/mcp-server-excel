# Command Reference

Complete reference for all ExcelMcp.CLI commands.

## File Commands

Create and manage Excel workbooks.

**create-empty** - Create empty Excel workbook

```powershell
excelcli create-empty <file.xlsx|file.xlsm>
```

Essential for coding agents to programmatically create Excel workbooks. Use `.xlsm` extension for macro-enabled workbooks that support VBA.

## Power Query Commands (`pq-*`)

Manage Power Query queries in Excel workbooks.

**pq-list** - List all Power Queries

```powershell
excelcli pq-list <file.xlsx>
```

**pq-view** - View Power Query M code

```powershell
excelcli pq-view <file.xlsx> <query-name>
```

**pq-import** - Create or import query from file

```powershell
excelcli pq-import <file.xlsx> <query-name> <source.pq>
```

**pq-export** - Export query to file

```powershell
excelcli pq-export <file.xlsx> <query-name> <output.pq>
```

**pq-update** - Update existing query from file

```powershell
excelcli pq-update <file.xlsx> <query-name> <code.pq>
```

**pq-refresh** - Refresh query data

```powershell
excelcli pq-refresh <file.xlsx> <query-name>
```

**pq-loadto** - Load connection-only query to worksheet

```powershell
excelcli pq-loadto <file.xlsx> <query-name> <sheet-name>
```

**pq-delete** - Delete Power Query

```powershell
excelcli pq-delete <file.xlsx> <query-name>
```

## Sheet Commands (`sheet-*`)

Manage worksheets in Excel workbooks.

**sheet-list** - List all worksheets

```powershell
excelcli sheet-list <file.xlsx>
```

**sheet-read** - Read data from worksheet

```powershell
excelcli sheet-read <file.xlsx> <sheet-name> [range]

# Examples
excelcli sheet-read "Plan.xlsx" "Data"           # Read entire used range
excelcli sheet-read "Plan.xlsx" "Data" "A1:C10"  # Read specific range
```

**sheet-write** - Write CSV data to worksheet

```powershell
excelcli sheet-write <file.xlsx> <sheet-name> <data.csv>
```

**sheet-create** - Create new worksheet

```powershell
excelcli sheet-create <file.xlsx> <sheet-name>
```

**sheet-copy** - Copy worksheet

```powershell
excelcli sheet-copy <file.xlsx> <source-sheet> <new-sheet>
```

**sheet-rename** - Rename worksheet

```powershell
excelcli sheet-rename <file.xlsx> <old-name> <new-name>
```

**sheet-delete** - Delete worksheet

```powershell
excelcli sheet-delete <file.xlsx> <sheet-name>
```

**sheet-clear** - Clear worksheet data

```powershell
excelcli sheet-clear <file.xlsx> <sheet-name> [range]
```

**sheet-append** - Append CSV data to worksheet

```powershell
excelcli sheet-append <file.xlsx> <sheet-name> <data.csv>
```

## Parameter Commands (`param-*`)

Manage named ranges and parameters.

**param-list** - List all named ranges

```powershell
excelcli param-list <file.xlsx>
```

**param-get** - Get named range value

```powershell
excelcli param-get <file.xlsx> <param-name>
```

**param-set** - Set named range value

```powershell
excelcli param-set <file.xlsx> <param-name> <value>
```

**param-create** - Create named range

```powershell
excelcli param-create <file.xlsx> <param-name> <reference>
```

**param-delete** - Delete named range

```powershell
excelcli param-delete <file.xlsx> <param-name>
```

## Cell Commands (`cell-*`)

Manage individual cells.

**cell-get-value** - Get cell value

```powershell
excelcli cell-get-value <file.xlsx> <sheet-name> <cell>
```

**cell-set-value** - Set cell value

```powershell
excelcli cell-set-value <file.xlsx> <sheet-name> <cell> <value>
```

**cell-get-formula** - Get cell formula

```powershell
excelcli cell-get-formula <file.xlsx> <sheet-name> <cell>
```

**cell-set-formula** - Set cell formula

```powershell
excelcli cell-set-formula <file.xlsx> <sheet-name> <cell> <formula>
```

## VBA Script Commands (`script-*`)

**⚠️ VBA commands require macro-enabled (.xlsm) files!**

Manage VBA scripts and macros in macro-enabled Excel workbooks.

**script-list** - List all VBA modules and procedures

```powershell
excelcli script-list <file.xlsm>
```

**script-export** - Export VBA module to file

```powershell
excelcli script-export <file.xlsm> <module-name> <output.vba>
```

**script-import** - Import VBA module from file

```powershell
excelcli script-import <file.xlsm> <module-name> <source.vba>
```

**script-update** - Update existing VBA module

```powershell
excelcli script-update <file.xlsm> <module-name> <source.vba>
```

**script-run** - Execute VBA macro with parameters

```powershell
excelcli script-run <file.xlsm> <macro-name> [param1] [param2] ...

# Examples
excelcli script-run "Report.xlsm" "ProcessData"
excelcli script-run "Analysis.xlsm" "CalculateTotal" "Sheet1" "A1:C10"
```

**script-delete** - Remove VBA module

```powershell
excelcli script-delete <file.xlsm> <module-name>
```

## Setup Commands

Configure VBA trust settings for automation.

**setup-vba-trust** - Enable VBA project access

```powershell
excelcli setup-vba-trust
```

**check-vba-trust** - Check VBA trust configuration

```powershell
excelcli check-vba-trust
```

## File Format Support

- **`.xlsx`** - Standard Excel workbooks (Power Query, worksheets, parameters)
- **`.xlsm`** - Macro-enabled workbooks (includes VBA script support)

Use `create-empty` with `.xlsm` extension to create macro-enabled workbooks:

```powershell
excelcli create-empty "macros.xlsm"  # Creates macro-enabled workbook
excelcli create-empty "data.xlsx"    # Creates standard workbook
```

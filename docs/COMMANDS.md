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
excelcli pq-import <file.xlsx> <query-name> <source.pq> [--privacy-level <None|Private|Organizational|Public>]
```

Import a Power Query from an M code file. If the query combines data from multiple sources and privacy level is not specified, you'll receive guidance on which privacy level to choose.

**pq-export** - Export query to file

```powershell
excelcli pq-export <file.xlsx> <query-name> <output.pq>
```

**pq-update** - Update existing query from file

```powershell
excelcli pq-update <file.xlsx> <query-name> <code.pq> [--privacy-level <None|Private|Organizational|Public>]
```

Update an existing Power Query with new M code. If the query combines data from multiple sources and privacy level is not specified, you'll receive guidance on which privacy level to choose.

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

## VBA Trust Configuration

VBA operations require **"Trust access to the VBA project object model"** to be enabled in Excel settings. This is a one-time manual setup for security reasons.

### How to Enable VBA Trust

1. Open Excel
2. Go to **File → Options → Trust Center**
3. Click **"Trust Center Settings"**
4. Select **"Macro Settings"**
5. Check **"✓ Trust access to the VBA project object model"**
6. Click **OK** twice to save settings

After enabling this setting, VBA operations will work automatically. If VBA trust is not enabled, commands will display detailed instructions.

**Security Note:** ExcelMcp never modifies security settings automatically. Users must explicitly enable VBA trust through Excel's settings to maintain security control.

For more information, see [Microsoft's documentation on macro security](https://support.microsoft.com/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6).

## Power Query Privacy Levels

When Power Query combines data from multiple sources, Excel requires a privacy level to be specified for security. ExcelMcp provides explicit control through the `--privacy-level` parameter.

### Privacy Level Options

- **None** - Ignores privacy levels, allows combining any data sources (least secure)
- **Private** - Prevents sharing data with other sources (most secure, recommended for sensitive data)
- **Organizational** - Data can be shared within organization (recommended for internal data)
- **Public** - Publicly available data sources (appropriate for public APIs)

### Using Privacy Levels

```powershell
# Specify privacy level explicitly
excelcli pq-import data.xlsx "WebData" query.pq --privacy-level Private

# Set default via environment variable (useful for automation)
$env:EXCEL_DEFAULT_PRIVACY_LEVEL = "Private"
excelcli pq-import data.xlsx "WebData" query.pq
```

If a privacy level is needed but not specified, the command will display:
- Existing privacy levels in the workbook
- Recommended privacy level based on your queries
- Clear instructions on how to proceed

**Security Note:** ExcelMcp never applies privacy levels automatically. Users must explicitly choose the appropriate level for their data security requirements.

## File Format Support

- **`.xlsx`** - Standard Excel workbooks (Power Query, worksheets, parameters)
- **`.xlsm`** - Macro-enabled workbooks (includes VBA script support)

Use `create-empty` with `.xlsm` extension to create macro-enabled workbooks:

```powershell
excelcli create-empty "macros.xlsm"  # Creates macro-enabled workbook
excelcli create-empty "data.xlsx"    # Creates standard workbook
```

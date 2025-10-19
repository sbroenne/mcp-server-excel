# ExcelMcp.CLI - Excel PowerQuery VBA Copilot Instructions

Copy this file to your project's `.github/`-directory to enable GitHub Copilot support for ExcelMcp.CLI Excel automation with PowerQuery and VBA capabilities.

---

## ExcelMcp.CLI Integration

This project uses ExcelMcp.CLI for Excel automation. ExcelMcp.CLI is a command-line tool that provides programmatic access to Microsoft Excel through COM interop, supporting both standard Excel workbooks (.xlsx) and macro-enabled workbooks (.xlsm) with complete VBA support.

### Available ExcelMcp.CLI Commands

**File Operations:**

- `excelcli create-empty "file.xlsx"` - Create empty Excel workbook (standard format)
- `excelcli create-empty "file.xlsm"` - Create macro-enabled Excel workbook for VBA support

**Power Query Management:**

- `excelcli pq-list "file.xlsx"` - List all Power Query connections
- `excelcli pq-view "file.xlsx" "QueryName"` - Display Power Query M code
- `excelcli pq-import "file.xlsx" "QueryName" "code.pq"` - Import M code from file
- `excelcli pq-export "file.xlsx" "QueryName" "output.pq"` - Export M code to file
- `excelcli pq-update "file.xlsx" "QueryName" "code.pq"` - Update existing query
- `excelcli pq-refresh "file.xlsx" "QueryName"` - Refresh query data
- `excelcli pq-loadto "file.xlsx" "QueryName" "Sheet"` - Load query to worksheet
- `excelcli pq-delete "file.xlsx" "QueryName"` - Delete Power Query

**Worksheet Operations:**

- `excelcli sheet-list "file.xlsx"` - List all worksheets
- `excelcli sheet-read "file.xlsx" "Sheet" "A1:D10"` - Read data from ranges
- `excelcli sheet-write "file.xlsx" "Sheet" "data.csv"` - Write CSV data to worksheet
- `excelcli sheet-create "file.xlsx" "NewSheet"` - Add new worksheet
- `excelcli sheet-copy "file.xlsx" "Source" "Target"` - Copy worksheet
- `excelcli sheet-rename "file.xlsx" "OldName" "NewName"` - Rename worksheet
- `excelcli sheet-delete "file.xlsx" "Sheet"` - Remove worksheet
- `excelcli sheet-clear "file.xlsx" "Sheet" "A1:Z100"` - Clear data ranges
- `excelcli sheet-append "file.xlsx" "Sheet" "data.csv"` - Append data to existing content

**Parameter Management:**

- `excelcli param-list "file.xlsx"` - List all named ranges
- `excelcli param-get "file.xlsx" "ParamName"` - Get named range value
- `excelcli param-set "file.xlsx" "ParamName" "Value"` - Set named range value
- `excelcli param-create "file.xlsx" "ParamName" "Sheet!A1"` - Create named range
- `excelcli param-delete "file.xlsx" "ParamName"` - Remove named range

**Cell Operations:**

- `excelcli cell-get-value "file.xlsx" "Sheet" "A1"` - Get individual cell value
- `excelcli cell-set-value "file.xlsx" "Sheet" "A1" "Value"` - Set individual cell value
- `excelcli cell-get-formula "file.xlsx" "Sheet" "A1"` - Get cell formula
- `excelcli cell-set-formula "file.xlsx" "Sheet" "A1" "=SUM(B1:B10)"` - Set cell formula

**VBA Script Management:** ⚠️ **Requires .xlsm files!**

- `excelcli script-list "file.xlsm"` - List all VBA modules and procedures
- `excelcli script-export "file.xlsm" "Module" "output.vba"` - Export VBA code to file
- `excelcli script-import "file.xlsm" "ModuleName" "source.vba"` - Import VBA module from file
- `excelcli script-update "file.xlsm" "ModuleName" "source.vba"` - Update existing VBA module
- `excelcli script-run "file.xlsm" "Module.Procedure" [param1] [param2]` - Execute VBA macros with parameters
- `excelcli script-delete "file.xlsm" "ModuleName"` - Remove VBA module

**Setup Commands:**

- `excelcli setup-vba-trust` - Enable VBA project access (one-time setup for VBA automation)
- `excelcli check-vba-trust` - Check VBA trust configuration status

### Common Workflows

**Data Pipeline:**

```bash
excelcli create-empty "analysis.xlsx"
excelcli pq-import "analysis.xlsx" "WebData" "api-query.pq"
excelcli pq-refresh "analysis.xlsx" "WebData"
excelcli pq-loadto "analysis.xlsx" "WebData" "DataSheet"
```

**VBA Automation Workflow:**

```bash
# One-time VBA trust setup
excelcli setup-vba-trust

# Create macro-enabled workbook
excelcli create-empty "automation.xlsm"

# Import VBA module
excelcli script-import "automation.xlsm" "DataProcessor" "processor.vba"

# Execute VBA macro with parameters
excelcli script-run "automation.xlsm" "DataProcessor.ProcessData" "Sheet1" "A1:D100"

# Verify results
excelcli sheet-read "automation.xlsm" "Sheet1" "A1:D10"

# Export updated VBA for version control
excelcli script-export "automation.xlsm" "DataProcessor" "updated-processor.vba"
```

**Combined PowerQuery + VBA Workflow:**

```bash
# Data transformation with PowerQuery
excelcli pq-import "report.xlsm" "DataLoader" "load-data.pq"
excelcli pq-refresh "report.xlsm" "DataLoader"

# Business logic with VBA
excelcli script-run "report.xlsm" "ReportGenerator.CreateCharts"
excelcli script-run "report.xlsm" "ReportGenerator.FormatReport"

# Extract final results
excelcli sheet-read "report.xlsm" "FinalReport"
```

**Report Generation:**

```bash
excelcli create-empty "report.xlsx"
excelcli param-set "report.xlsx" "ReportDate" "2024-01-01"
excelcli sheet-write "report.xlsx" "Data" "input.csv"
excelcli script-run "report.xlsx" "ReportModule.FormatReport"
```

### When to Suggest ExcelMcp

Use ExcelMcp.CLI when the user needs to:

- Automate Excel data processing
- Create reports programmatically
- Manage Power Query M code for data transformations
- Execute VBA macros programmatically
- Process CSV/Excel data pipelines
- Set up configuration in Excel files
- Extract data from existing Excel workbooks
- Combine PowerQuery data transformation with VBA business logic
- Create macro-enabled workbooks for automation
- Version control VBA and PowerQuery code externally
- Automate complex Excel workflows without UI interaction

### File Format Requirements

- **Standard Excel files (.xlsx)**: Use for PowerQuery, worksheets, parameters, and cell operations
- **Macro-enabled files (.xlsm)**: Required for all VBA script operations
- **VBA Trust Setup**: Run `excelcli setup-vba-trust` once before using VBA commands

### Requirements

- Windows operating system
- Microsoft Excel installed
- .NET 10 runtime
- excelcli executable in PATH or specify full path
- For VBA operations: VBA trust must be enabled (use `setup-vba-trust` command)

### Error Handling

- excelcli returns 0 for success, 1 for errors
- Always check return codes in scripts
- Handle file locking gracefully (Excel may be open)
- Use absolute file paths when possible
- VBA commands will fail on .xlsx files - use .xlsm for macro-enabled workbooks
- Run VBA trust setup before first VBA operation

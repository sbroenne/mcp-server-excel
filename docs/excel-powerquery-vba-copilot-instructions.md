# ExcelMcp - Excel PowerQuery VBA Copilot Instructions

Copy this file to your project's `.github/`-directory to enable GitHub Copilot support for ExcelMcp Excel automation with PowerQuery and VBA capabilities.

---

## ExcelMcp Integration

This project uses ExcelMcp for Excel automation. ExcelMcp is a command-line tool that provides programmatic access to Microsoft Excel through COM interop, supporting both standard Excel workbooks (.xlsx) and macro-enabled workbooks (.xlsm) with complete VBA support.

### Available ExcelMcp Commands

**File Operations:**

- `ExcelMcp create-empty "file.xlsx"` - Create empty Excel workbook (standard format)
- `ExcelMcp create-empty "file.xlsm"` - Create macro-enabled Excel workbook for VBA support

**Power Query Management:**

- `ExcelMcp pq-list "file.xlsx"` - List all Power Query connections
- `ExcelMcp pq-view "file.xlsx" "QueryName"` - Display Power Query M code
- `ExcelMcp pq-import "file.xlsx" "QueryName" "code.pq"` - Import M code from file
- `ExcelMcp pq-export "file.xlsx" "QueryName" "output.pq"` - Export M code to file
- `ExcelMcp pq-update "file.xlsx" "QueryName" "code.pq"` - Update existing query
- `ExcelMcp pq-refresh "file.xlsx" "QueryName"` - Refresh query data
- `ExcelMcp pq-loadto "file.xlsx" "QueryName" "Sheet"` - Load query to worksheet
- `ExcelMcp pq-delete "file.xlsx" "QueryName"` - Delete Power Query

**Worksheet Operations:**

- `ExcelMcp sheet-list "file.xlsx"` - List all worksheets
- `ExcelMcp sheet-read "file.xlsx" "Sheet" "A1:D10"` - Read data from ranges
- `ExcelMcp sheet-write "file.xlsx" "Sheet" "data.csv"` - Write CSV data to worksheet
- `ExcelMcp sheet-create "file.xlsx" "NewSheet"` - Add new worksheet
- `ExcelMcp sheet-copy "file.xlsx" "Source" "Target"` - Copy worksheet
- `ExcelMcp sheet-rename "file.xlsx" "OldName" "NewName"` - Rename worksheet
- `ExcelMcp sheet-delete "file.xlsx" "Sheet"` - Remove worksheet
- `ExcelMcp sheet-clear "file.xlsx" "Sheet" "A1:Z100"` - Clear data ranges
- `ExcelMcp sheet-append "file.xlsx" "Sheet" "data.csv"` - Append data to existing content

**Parameter Management:**

- `ExcelMcp param-list "file.xlsx"` - List all named ranges
- `ExcelMcp param-get "file.xlsx" "ParamName"` - Get named range value
- `ExcelMcp param-set "file.xlsx" "ParamName" "Value"` - Set named range value
- `ExcelMcp param-create "file.xlsx" "ParamName" "Sheet!A1"` - Create named range
- `ExcelMcp param-delete "file.xlsx" "ParamName"` - Remove named range

**Cell Operations:**

- `ExcelMcp cell-get-value "file.xlsx" "Sheet" "A1"` - Get individual cell value
- `ExcelMcp cell-set-value "file.xlsx" "Sheet" "A1" "Value"` - Set individual cell value
- `ExcelMcp cell-get-formula "file.xlsx" "Sheet" "A1"` - Get cell formula
- `ExcelMcp cell-set-formula "file.xlsx" "Sheet" "A1" "=SUM(B1:B10)"` - Set cell formula

**VBA Script Management:** ⚠️ **Requires .xlsm files!**

- `ExcelMcp script-list "file.xlsm"` - List all VBA modules and procedures
- `ExcelMcp script-export "file.xlsm" "Module" "output.vba"` - Export VBA code to file
- `ExcelMcp script-import "file.xlsm" "ModuleName" "source.vba"` - Import VBA module from file
- `ExcelMcp script-update "file.xlsm" "ModuleName" "source.vba"` - Update existing VBA module
- `ExcelMcp script-run "file.xlsm" "Module.Procedure" [param1] [param2]` - Execute VBA macros with parameters
- `ExcelMcp script-delete "file.xlsm" "ModuleName"` - Remove VBA module

**Setup Commands:**

- `ExcelMcp setup-vba-trust` - Enable VBA project access (one-time setup for VBA automation)
- `ExcelMcp check-vba-trust` - Check VBA trust configuration status

### Common Workflows

**Data Pipeline:**

```bash
ExcelMcp create-empty "analysis.xlsx"
ExcelMcp pq-import "analysis.xlsx" "WebData" "api-query.pq"
ExcelMcp pq-refresh "analysis.xlsx" "WebData"
ExcelMcp pq-loadto "analysis.xlsx" "WebData" "DataSheet"
```

**VBA Automation Workflow:**

```bash
# One-time VBA trust setup
ExcelMcp setup-vba-trust

# Create macro-enabled workbook
ExcelMcp create-empty "automation.xlsm"

# Import VBA module
ExcelMcp script-import "automation.xlsm" "DataProcessor" "processor.vba"

# Execute VBA macro with parameters
ExcelMcp script-run "automation.xlsm" "DataProcessor.ProcessData" "Sheet1" "A1:D100"

# Verify results
ExcelMcp sheet-read "automation.xlsm" "Sheet1" "A1:D10"

# Export updated VBA for version control
ExcelMcp script-export "automation.xlsm" "DataProcessor" "updated-processor.vba"
```

**Combined PowerQuery + VBA Workflow:**

```bash
# Data transformation with PowerQuery
ExcelMcp pq-import "report.xlsm" "DataLoader" "load-data.pq"
ExcelMcp pq-refresh "report.xlsm" "DataLoader"

# Business logic with VBA
ExcelMcp script-run "report.xlsm" "ReportGenerator.CreateCharts"
ExcelMcp script-run "report.xlsm" "ReportGenerator.FormatReport"

# Extract final results
ExcelMcp sheet-read "report.xlsm" "FinalReport"
```

**Report Generation:**

```bash
ExcelMcp create-empty "report.xlsx"
ExcelMcp param-set "report.xlsx" "ReportDate" "2024-01-01"
ExcelMcp sheet-write "report.xlsx" "Data" "input.csv"
ExcelMcp script-run "report.xlsx" "ReportModule.FormatReport"
```

### When to Suggest ExcelMcp

Use ExcelMcp when the user needs to:

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
- **VBA Trust Setup**: Run `ExcelMcp setup-vba-trust` once before using VBA commands

### Requirements

- Windows operating system
- Microsoft Excel installed
- .NET 8.0 runtime
- ExcelMcp executable in PATH or specify full path
- For VBA operations: VBA trust must be enabled (use `setup-vba-trust` command)

### Error Handling

- ExcelMcp returns 0 for success, 1 for errors
- Always check return codes in scripts
- Handle file locking gracefully (Excel may be open)
- Use absolute file paths when possible
- VBA commands will fail on .xlsx files - use .xlsm for macro-enabled workbooks
- Run VBA trust setup before first VBA operation

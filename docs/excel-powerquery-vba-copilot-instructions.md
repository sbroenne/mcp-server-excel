# ExcelMcp.CLI - Excel PowerQuery VBA Copilot Instructions

Copy this file to your project's `.github/`-directory to enable GitHub Copilot support for ExcelMcp.CLI Excel automation with PowerQuery and VBA capabilities.

---

## ExcelMcp.CLI Integration

This project uses ExcelMcp.CLI for Excel automation. ExcelMcp.CLI is a command-line tool that provides programmatic access to Microsoft Excel through COM interop, supporting both standard Excel workbooks (.xlsx) and macro-enabled workbooks (.xlsm) with complete VBA support.

### Available ExcelMcp.CLI Commands

**File Operations:**

- `ExcelMcp.CLI create-empty "file.xlsx"` - Create empty Excel workbook (standard format)
- `ExcelMcp.CLI create-empty "file.xlsm"` - Create macro-enabled Excel workbook for VBA support

**Power Query Management:**

- `ExcelMcp.CLI pq-list "file.xlsx"` - List all Power Query connections
- `ExcelMcp.CLI pq-view "file.xlsx" "QueryName"` - Display Power Query M code
- `ExcelMcp.CLI pq-import "file.xlsx" "QueryName" "code.pq"` - Import M code from file
- `ExcelMcp.CLI pq-export "file.xlsx" "QueryName" "output.pq"` - Export M code to file
- `ExcelMcp.CLI pq-update "file.xlsx" "QueryName" "code.pq"` - Update existing query
- `ExcelMcp.CLI pq-refresh "file.xlsx" "QueryName"` - Refresh query data
- `ExcelMcp.CLI pq-loadto "file.xlsx" "QueryName" "Sheet"` - Load query to worksheet
- `ExcelMcp.CLI pq-delete "file.xlsx" "QueryName"` - Delete Power Query

**Worksheet Operations:**

- `ExcelMcp.CLI sheet-list "file.xlsx"` - List all worksheets
- `ExcelMcp.CLI sheet-read "file.xlsx" "Sheet" "A1:D10"` - Read data from ranges
- `ExcelMcp.CLI sheet-write "file.xlsx" "Sheet" "data.csv"` - Write CSV data to worksheet
- `ExcelMcp.CLI sheet-create "file.xlsx" "NewSheet"` - Add new worksheet
- `ExcelMcp.CLI sheet-copy "file.xlsx" "Source" "Target"` - Copy worksheet
- `ExcelMcp.CLI sheet-rename "file.xlsx" "OldName" "NewName"` - Rename worksheet
- `ExcelMcp.CLI sheet-delete "file.xlsx" "Sheet"` - Remove worksheet
- `ExcelMcp.CLI sheet-clear "file.xlsx" "Sheet" "A1:Z100"` - Clear data ranges
- `ExcelMcp.CLI sheet-append "file.xlsx" "Sheet" "data.csv"` - Append data to existing content

**Parameter Management:**

- `ExcelMcp.CLI param-list "file.xlsx"` - List all named ranges
- `ExcelMcp.CLI param-get "file.xlsx" "ParamName"` - Get named range value
- `ExcelMcp.CLI param-set "file.xlsx" "ParamName" "Value"` - Set named range value
- `ExcelMcp.CLI param-create "file.xlsx" "ParamName" "Sheet!A1"` - Create named range
- `ExcelMcp.CLI param-delete "file.xlsx" "ParamName"` - Remove named range

**Cell Operations:**

- `ExcelMcp.CLI cell-get-value "file.xlsx" "Sheet" "A1"` - Get individual cell value
- `ExcelMcp.CLI cell-set-value "file.xlsx" "Sheet" "A1" "Value"` - Set individual cell value
- `ExcelMcp.CLI cell-get-formula "file.xlsx" "Sheet" "A1"` - Get cell formula
- `ExcelMcp.CLI cell-set-formula "file.xlsx" "Sheet" "A1" "=SUM(B1:B10)"` - Set cell formula

**VBA Script Management:** ⚠️ **Requires .xlsm files!**

- `ExcelMcp.CLI script-list "file.xlsm"` - List all VBA modules and procedures
- `ExcelMcp.CLI script-export "file.xlsm" "Module" "output.vba"` - Export VBA code to file
- `ExcelMcp.CLI script-import "file.xlsm" "ModuleName" "source.vba"` - Import VBA module from file
- `ExcelMcp.CLI script-update "file.xlsm" "ModuleName" "source.vba"` - Update existing VBA module
- `ExcelMcp.CLI script-run "file.xlsm" "Module.Procedure" [param1] [param2]` - Execute VBA macros with parameters
- `ExcelMcp.CLI script-delete "file.xlsm" "ModuleName"` - Remove VBA module

**Setup Commands:**

- `ExcelMcp.CLI setup-vba-trust` - Enable VBA project access (one-time setup for VBA automation)
- `ExcelMcp.CLI check-vba-trust` - Check VBA trust configuration status

### Common Workflows

**Data Pipeline:**

```bash
ExcelMcp.CLI create-empty "analysis.xlsx"
ExcelMcp.CLI pq-import "analysis.xlsx" "WebData" "api-query.pq"
ExcelMcp.CLI pq-refresh "analysis.xlsx" "WebData"
ExcelMcp.CLI pq-loadto "analysis.xlsx" "WebData" "DataSheet"
```

**VBA Automation Workflow:**

```bash
# One-time VBA trust setup
ExcelMcp.CLI setup-vba-trust

# Create macro-enabled workbook
ExcelMcp.CLI create-empty "automation.xlsm"

# Import VBA module
ExcelMcp.CLI script-import "automation.xlsm" "DataProcessor" "processor.vba"

# Execute VBA macro with parameters
ExcelMcp.CLI script-run "automation.xlsm" "DataProcessor.ProcessData" "Sheet1" "A1:D100"

# Verify results
ExcelMcp.CLI sheet-read "automation.xlsm" "Sheet1" "A1:D10"

# Export updated VBA for version control
ExcelMcp.CLI script-export "automation.xlsm" "DataProcessor" "updated-processor.vba"
```

**Combined PowerQuery + VBA Workflow:**

```bash
# Data transformation with PowerQuery
ExcelMcp.CLI pq-import "report.xlsm" "DataLoader" "load-data.pq"
ExcelMcp.CLI pq-refresh "report.xlsm" "DataLoader"

# Business logic with VBA
ExcelMcp.CLI script-run "report.xlsm" "ReportGenerator.CreateCharts"
ExcelMcp.CLI script-run "report.xlsm" "ReportGenerator.FormatReport"

# Extract final results
ExcelMcp.CLI sheet-read "report.xlsm" "FinalReport"
```

**Report Generation:**

```bash
ExcelMcp.CLI create-empty "report.xlsx"
ExcelMcp.CLI param-set "report.xlsx" "ReportDate" "2024-01-01"
ExcelMcp.CLI sheet-write "report.xlsx" "Data" "input.csv"
ExcelMcp.CLI script-run "report.xlsx" "ReportModule.FormatReport"
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
- **VBA Trust Setup**: Run `ExcelMcp.CLI setup-vba-trust` once before using VBA commands

### Requirements

- Windows operating system
- Microsoft Excel installed
- .NET 10 runtime
- ExcelMcp.CLI executable in PATH or specify full path
- For VBA operations: VBA trust must be enabled (use `setup-vba-trust` command)

### Error Handling

- ExcelMcp.CLI returns 0 for success, 1 for errors
- Always check return codes in scripts
- Handle file locking gracefully (Excel may be open)
- Use absolute file paths when possible
- VBA commands will fail on .xlsx files - use .xlsm for macro-enabled workbooks
- Run VBA trust setup before first VBA operation

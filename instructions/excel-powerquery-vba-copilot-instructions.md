# Excel MCP Server - AI-Assisted Excel Development Guide

Copy this file to your project's `.github/copilot-instructions.md` to enable GitHub Copilot support for ExcelMcp MCP Server - AI-powered Excel development with PowerQuery and VBA capabilities.

---

## Overview

The Excel MCP Server (ExcelMcp) provides AI-powered Excel development capabilities through Model Context Protocol (MCP) integration with GitHub Copilot. This enables conversational, natural language-driven Excel automation for Power Query, VBA, worksheets, named ranges, and more.

## What is the Excel MCP Server?

The Excel MCP Server is a bridge between GitHub Copilot and Microsoft Excel, allowing you to:

- **Develop Power Query M code** with AI assistance
- **Manage VBA scripts** through conversational commands
- **Manipulate worksheets and data** programmatically
- **Configure named ranges** as parameters
- **Review and optimize** existing Excel code
- **Version control** Power Query and VBA code

## Quick Start

**Want to get started immediately? Just ask Copilot:**

```
Set up my workspace for Excel MCP Server development. Check if .NET SDK 10 is installed, install it if needed, then configure the Excel MCP Server.
```

Copilot will automatically:

1. Check if .NET SDK 10 is installed
2. Install .NET SDK 10 via winget if not present
3. Verify the installation
4. Create `.vscode/mcp.json` with Excel MCP Server configuration
5. Confirm everything is ready to use

**Then verify the setup:**

```
List all available Excel MCP tools
```

You should see 6 Excel resources (excel_file, excel_powerquery, excel_worksheet, excel_parameter, excel_cell, excel_vba).

---

## Detailed Setup and Configuration

If you prefer to understand each step or need to troubleshoot, follow the detailed instructions below.

### Prerequisites

Before setting up the Excel MCP Server, ensure you have:

- **Windows OS** (required for Excel COM automation)
- **Microsoft Excel** installed (Desktop version)
- **.NET SDK** (version 10.0 or higher) - See installation below
- **VS Code** with GitHub Copilot extension

### Installing Prerequisites

#### Install .NET SDK

**Ask Copilot to install it for you:**

```
Install .NET SDK 10 using winget
```

**Or run manually in PowerShell:**

```powershell
winget install Microsoft.DotNet.SDK.10
```

After installation, verify it's installed:

```powershell
dotnet --version
```

You should see version 10.0.0 or higher.

### MCP Server Installation

#### Simple Setup: Just Ask Copilot! (Recommended)

The easiest way to get started is to simply ask GitHub Copilot to configure the MCP server for you.

**For New MCP Configuration:**

```
Set up the Excel MCP Server in this workspace
```

**For Existing MCP Configuration:**

If you already have other MCP servers (like browser, upstash, etc.), Copilot will automatically detect them and add Excel without removing your existing servers:

```
Add the Excel MCP Server to my existing MCP configuration
```

Copilot will automatically:

- Create `.vscode/` directory if needed
- Create or update `mcp.json` with the correct configuration
- Preserve any existing MCP servers
- Verify the setup is correct

#### Manual Setup (Alternative)

If you prefer to configure manually, create or update `.vscode/mcp.json` in your workspace root:

**New Configuration:**

```json
{
  "servers": {
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer@latest", "--yes"]
    }
  }
}
```

**Adding to Existing Servers:**

```json
{
  "servers": {
    "browser": {
      "command": "...",
      "args": ["..."]
    },
    "excel": {
      "command": "dnx",
      "args": ["Sbroenne.ExcelMcp.McpServer@latest", "--yes"]
    }
  }
}
```

### Verification

1. Open VS Code with your workspace containing `.vscode/mcp.json`
2. Open GitHub Copilot Chat
3. Ask: "List all Excel MCP tools available"
4. You should see 6 main Excel resources listed

## Available MCP Tools

The Excel MCP Server provides 6 resource-based tools for Excel automation:

### 1. excel_file - File Management

**Actions:** `create-empty`

**Use Cases:**

- Create new Excel workbooks (.xlsx)
- Create macro-enabled workbooks (.xlsm)

**Example Prompts:**

- "Create a new Excel workbook at C:\reports\analysis.xlsx"
- "Create a macro-enabled workbook for automation scripts"

### 2. excel_powerquery - Power Query Operations

**Actions:** `list`, `view`, `import`, `export`, `update`, `refresh`, `delete`, `set-load-to-table`, `set-load-to-data-model`, `set-load-to-both`, `set-connection-only`, `get-load-config`

**Use Cases:**

- List all Power Query queries in a workbook
- View Power Query M code
- Import queries from .pq files (version control)
- Export queries to .pq files
- Update existing queries with optimized code
- Refresh query data
- Configure query load destinations
- Delete obsolete queries

**Example Prompts:**

- "List all Power Query queries in analysis.xlsx"
- "View the M code for the 'SalesData' query"
- "Import the CustomerTransform.pq file into this workbook"
- "Export all Power Query queries to .pq files for version control"
- "Refresh the 'MonthlyReport' query"
- "Set the 'WebData' query to load to a worksheet table"
- "Review this Power Query and optimize for performance"

### 3. excel_worksheet - Worksheet Operations

**Actions:** `list`, `read`, `write`, `create`, `rename`, `copy`, `delete`, `clear`, `append`

**Use Cases:**

- List all worksheets in a workbook
- Read data from specific ranges
- Write data to worksheets (from CSV or arrays)
- Create new worksheets
- Rename, copy, or delete worksheets
- Clear worksheet data
- Append data to existing tables

**Example Prompts:**

- "List all worksheets in the workbook"
- "Read data from Sheet1 range A1:D10"
- "Create a new worksheet called 'Summary'"
- "Write the CSV data to the 'Data' worksheet"
- "Clear all data from the 'Temp' worksheet"
- "Copy the 'Template' worksheet to 'Report_2025'"

### 4. excel_parameter - Named Range Management

**Actions:** `list`, `get`, `set`, `create`, `delete`

**Use Cases:**

- List all named ranges in a workbook
- Get parameter values from named ranges
- Set parameter values
- Create new named ranges
- Delete obsolete named ranges

**Example Prompts:**

- "List all named ranges in parameters.xlsx"
- "Get the value of the 'StartDate' parameter"
- "Set the 'ReportMonth' parameter to '2025-10'"
- "Create a named range 'TaxRate' pointing to Setup!B5"
- "Delete the obsolete 'OldParameter' named range"

### 5. excel_cell - Cell Operations

**Actions:** `get-value`, `set-value`, `get-formula`, `set-formula`

**Use Cases:**

- Get individual cell values
- Set cell values programmatically
- Retrieve cell formulas
- Update cell formulas

**Example Prompts:**

- "Get the value from cell B5 in the Setup sheet"
- "Set cell C10 to the value 1000"
- "Get the formula from cell D20"
- "Set cell E5 formula to =SUM(E1:E4)"

### 6. excel_vba - VBA Script Management

⚠️ **Requires .xlsm files!**

**Actions:** `list`, `export`, `import`, `update`, `run`, `delete`

**Use Cases:**

- List all VBA modules in a workbook
- Export VBA modules to .bas files (version control)
- Import VBA modules from files
- Update existing VBA code
- Execute VBA procedures with parameters
- Delete obsolete VBA modules

**Example Prompts:**

- "List all VBA modules in automation.xlsm"
- "Export the 'DataProcessor' module to version control"
- "Import the ErrorHandler.bas module into this workbook"
- "Run the ProcessData procedure from the Automation module"
- "Add comprehensive error handling to the VBA module"
- "Delete the obsolete 'LegacyCode' module"

## AI-Assisted Development Workflows

### Power Query Development

#### Scenario: Create and Optimize a Data Transformation Query

**Step 1: Ask Copilot to Create a Query**

```
Create a Power Query in report.xlsx that:
1. Loads data from the 'RawData' worksheet
2. Filters rows where Amount > 0
3. Groups by Category and sums the Amount
4. Sorts by total descending
```

**Step 2: Review and Test**

```
Refresh the query and show me the first 10 rows of output
```

**Step 3: Code Analysis and Optimization**

```
Review the 'DataTransform' query M code and suggest specific improvements I can make
```

*Note: I can view and help you understand M code, suggest structural improvements, and help you implement changes, but I don't have built-in M code performance analysis engines. I'll help you apply Power Query best practices.*

**Step 4: Version Control**

```
Export all Power Query queries to the 'powerquery/' folder as .pq files
```

### VBA Development and Enhancement

#### Scenario: Add Error Handling to Existing VBA Code

**Step 1: Review Existing Code**

```
List all VBA modules in automation.xlsm and show me the code for 'DataProcessor'
```

**Step 2: Request Enhancement**

```
Add comprehensive error handling to the DataProcessor module, including:
- Try-catch patterns for all external operations
- Logging errors to a worksheet
- User-friendly error messages
- Cleanup code in Finally blocks
```

**Step 3: Version Control**

```
Export the updated DataProcessor module to vba/DataProcessor.bas
```

### Combined Workflows

#### Scenario: Create a Complete Data Pipeline

**Full Pipeline Request:**

```
Set up a data pipeline in report.xlsm that:
1. Creates a Power Query to load data from a SharePoint CSV
2. Transforms the data (filter, aggregate, clean)
3. Loads results to a 'ProcessedData' worksheet
4. Creates a VBA macro to format the output and generate charts
5. Sets up named ranges for configuration (FilePath, ReportMonth)
```

Copilot will orchestrate all the MCP tools to build the complete solution.

## Best Practices

### File Management

**Always Close Excel Processes Before Operations:**

```powershell
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
```

This prevents file access errors during MCP operations.

**Use Macro-Enabled Format for VBA:**

- VBA operations require `.xlsm` files
- Standard Power Query/worksheet operations work with `.xlsx`

### Power Query Development

**Version Control Power Query M Code:**

1. Create a `powerquery/` directory in your project
2. Export queries to `.pq` files: "Export all queries to powerquery/"
3. Commit `.pq` files to Git
4. Import queries when setting up new workbooks

**Shared Sources Pattern:**

When multiple queries need the same data source (SharePoint, web, file):

```powerquery
// Shared source query (execute once, buffer result)
let
    Source = Table.Buffer(
        Excel.Workbook(
            Web.Contents("https://sharepoint.com/file.xlsx"),
            null,
            true
        )
    )
in
    Source
```

Reference the buffered source in dependent queries to avoid redundant downloads.

**Query Dependencies:**

- Create queries in dependency order (parameters → shared sources → transformations → output)
- Reference queries by name: `Source = SharedDataSource`
- Test incrementally after each query

### Named Ranges as Configuration

**Parameter Pattern:**

1. Create a "Setup" worksheet
2. Define parameters in A column, values in B column
3. Create named ranges pointing to value cells

**Example Setup:**

Ask Copilot:

```
Set up a parameter configuration in Setup worksheet:
- Cell A1: "Parameter", B1: "Value"
- Cell A2: "StartDate", B2: "2025-01-01"
- Cell A3: "EndDate", B3: "2025-12-31"
- Cell A4: "DataFilePath", B4: "C:\data\source.csv"
Create named ranges for each parameter pointing to column B
```

**Access in Power Query:**

```powerquery
StartDate = Excel.CurrentWorkbook(){[Name="StartDate"]}[Content]{0}[Value]
```

### VBA Development

**Error Handling Template:**

Ask Copilot to add this pattern:

```
Add error handling to all VBA procedures using this pattern:
- On Error GoTo ErrorHandler at the start
- Cleanup code before Exit Sub
- ErrorHandler label that logs errors and shows user-friendly messages
- Resume cleanup before exiting
```

**Version Control VBA:**

1. Export modules to `.bas` files: "Export all VBA modules to vba/"
2. Commit `.bas` files to Git
3. Import when setting up new workbooks

### Performance Optimization

**Power Query:**

- Ask: "Review this query for performance issues and query folding opportunities"
- Use `Table.Buffer()` on small lookup tables only
- Filter data early in transformation chain

**VBA:**

- Disable screen updating: `Application.ScreenUpdating = False`
- Use arrays for bulk operations
- Avoid cell-by-cell writes in loops

### Security Considerations

#### VBA Trust Settings

For VBA operations to work, Excel must trust VBA project access.

**Enable VBA Trust (via Copilot):**

```
Guide me through enabling VBA trust for MCP operations
```

#### Sensitive Data Handling

**Best Practices:**

- Never hardcode credentials in Power Query or VBA
- Use parameter files for sensitive configuration
- Exclude parameter files from version control (.gitignore)
- Use Windows Credential Manager for API keys

**Example .gitignore:**

```
# Sensitive configuration
**/Config_Local.xlsx
**/Secrets.xlsx
**/.env
```

#### Code Review Before Execution

**Always Review AI-Generated Code:**

```
Before running the VBA macro, show me the complete code and explain:
1. What it does step by step
2. Any potential risks
3. What data it accesses or modifies
```

## Common Workflows

### Workflow 1: Setting Up a New Report Workbook

**Single Prompt to Copilot:**

```
Create a new macro-enabled workbook 'MonthlyReport.xlsm' with:
1. Setup worksheet with parameters (ReportMonth, DataSource, OutputPath)
2. Named ranges for all parameters
3. Power Query to load data from parameter-specified CSV
4. Transform query to filter current month and aggregate by category
5. VBA macro to refresh queries and format output
6. Load query results to 'Report' worksheet
```

### Workflow 2: Migrating Power Query Between Workbooks

**Scenario:** Consolidate queries from multiple files into a master workbook

**Step 1: Export from Source Files**

```
Export all Power Query queries from source1.xlsx to temp/source1/
Export all Power Query queries from source2.xlsx to temp/source2/
```

**Step 2: Review and Organize**

```
List all .pq files in temp/ and show me their purposes
```

**Step 3: Import to Master Workbook**

```
Import these queries into master.xlsx in this order:
1. SharedSources.pq (foundation)
2. Parameters.pq
3. DataTransforms.pq
4. OutputQueries.pq
```

**Step 4: Create Supporting Structure**

```
Create named ranges that the imported queries reference:
- DataFilePath pointing to Setup!B2
- StartDate pointing to Setup!B3
```

**Step 5: Test**

```
Refresh all queries and verify no errors
```

### Workflow 3: Code Review and Optimization

**Power Query Review:**

```
Review all Power Query queries in analysis.xlsx. Show me the M code for each query and help me identify:
1. Potential structural improvements
2. Best practices I should apply
3. Common Power Query patterns I could use
4. Data loading configuration optimizations
Guide me through implementing the improvements.
```

*Note: I'll help you review and improve M code structure and apply Power Query best practices, but I don't have automated performance analysis engines. I provide guidance based on Power Query development patterns and your specific needs.*

**VBA Code Review:**

```
Review all VBA modules in automation.xlsm. Show me the code and help me identify:
1. Error handling gaps and help me add proper error handling
2. Code organization improvements
3. Security best practices I should apply
4. Performance optimizations I can implement
Guide me through applying the improvements.
```

### Workflow 4: Data Pipeline Automation

**Complete Pipeline Setup:**

```
Build an automated data pipeline:

1. Data Collection (Power Query):
   - Load from multiple CSV sources in 'data/' folder
   - Combine all files with same schema
   - Apply data type transformations

2. Data Processing (Power Query):
   - Filter invalid records
   - Aggregate by month and category
   - Join with lookup tables

3. Output Generation (VBA):
   - Refresh all queries
   - Create pivot tables
   - Generate summary charts
   - Export to PDF

4. Configuration (Named Ranges):
   - DataDirectory
   - ReportMonth
   - OutputPath

Set up the complete pipeline and create a 'Run Report' button to execute it
```

## Troubleshooting

### Common Errors

#### Error: File Access Denied

**Cause:** Excel process has file open

**Solution:**

```powershell
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
```

#### Error: Parameter Not Found

**Cause:** Named range doesn't exist or incorrect reference

**Solution:**

```
List all named ranges in the workbook
Verify the parameter name matches exactly (case-sensitive)
Create the missing named range if needed
```

#### Error: VBA Operation Failed (requires .xlsm)

**Cause:** Trying to use VBA operations on .xlsx file

**Solution:**

```
Save the file as .xlsm (macro-enabled format)
Re-run the VBA operation
```

#### Error: Circular Reference in Query

**Cause:** Named range or query references itself

**Solution:**

```
Review query dependencies
Ensure named ranges point to cells, not the file itself
Break circular reference chain
```

### Validation Steps

**After Query Import/Update:**

1. List all queries: "List all Power Query queries"
2. Check named ranges: "List all named ranges"
3. View query code: "View the M code for QueryName"
4. Test refresh: "Refresh the QueryName query"

**After VBA Import/Update:**

1. List modules: "List all VBA modules"
2. View code: "Show me the code for ModuleName"
3. Test execution: "Run the ProcedureName procedure"

## Advanced Patterns

### Dynamic Parameter Configuration

**Pattern:** Store parameters in Excel, read in both Power Query and VBA

**Setup:**

```
Create a parameters system:
1. Setup worksheet with key-value pairs
2. Named ranges for each parameter
3. Power Query helper function to read parameters
4. VBA function to access the same parameters
Document the parameter schema
```

### Incremental Data Loading

**Pattern:** Load only new/changed data to improve performance

**Implementation:**

```
Create an incremental load Power Query:
1. Store last refresh timestamp in a named range
2. Filter source data for records after last refresh
3. Append new data to existing table
4. Update refresh timestamp
5. Handle schema changes gracefully
```

### Data Validation and Quality Checks

**Pattern:** Automated data validation with error reporting

**Implementation:**

```
Build a data quality pipeline:
1. Power Query to load and validate data
2. Create validation rules (null checks, range checks, format checks)
3. Log validation errors to 'DataQuality' worksheet
4. VBA macro to generate quality report
5. Email notification if critical errors found
```

### Multi-Source Data Integration

**Pattern:** Combine data from Excel, CSV, SharePoint, and Web APIs

**Implementation:**

```
Create a multi-source data integration:
1. Shared buffered sources for each data source type
2. Schema normalization queries
3. Data merging and deduplication
4. Master output query
5. Refresh orchestration VBA macro
```

## Integration with Other Tools

### Git Version Control

**Recommended Structure:**

```
project/
├── .gitignore
├── README.md
├── workbooks/
│   ├── Report.xlsx
│   └── Analysis.xlsm
├── powerquery/
│   ├── shared/
│   │   ├── SharedSource.pq
│   │   └── Parameters.pq
│   ├── transforms/
│   │   ├── DataClean.pq
│   │   └── Aggregations.pq
│   └── output/
│       └── FinalReport.pq
└── vba/
    ├── DataProcessor.bas
    ├── ReportGenerator.bas
    └── ErrorHandler.bas
```

**Workflow:**

1. Export code: "Export all Power Query to powerquery/ and all VBA to vba/"
2. Commit to Git: `git add powerquery/ vba/`
3. Work in branch: `git checkout -b feature/new-report`
4. Import on other machines: "Import all .pq files from powerquery/ and all .bas files from vba/"

### Power BI Integration

**Pattern:** Use Excel as data preparation layer for Power BI

**Workflow:**

```
Set up Excel-to-Power BI pipeline:
1. Power Query in Excel for complex transformations
2. Export cleaned data to CSV or direct query
3. Power BI loads from Excel or exported files
4. Refresh orchestration via VBA or Power Automate
```

### Azure Integration

**Pattern:** Excel workbooks accessing Azure data sources

**Setup:**

```
Configure Azure data access:
1. Power Query to connect to Azure SQL, Blob Storage, or Data Lake
2. Named ranges for connection strings (reference Key Vault)
3. OAuth authentication configuration
4. Automatic token refresh handling
```

## Example Conversations with Copilot

### Example 1: Complete Report Setup

**User:** "I need to create a monthly sales report workbook that loads data from a CSV, aggregates by region and product, and generates charts."

**Copilot Actions:**

1. Creates new .xlsm workbook
2. Sets up Setup worksheet with parameters
3. Creates named ranges (ReportMonth, DataFilePath)
4. Imports Power Query to load CSV
5. Creates aggregation queries
6. Builds VBA macro for chart generation
7. Sets up refresh button

**User Follow-up:** "Add error handling to the VBA and optimize the Power Query for performance."

**Copilot Actions:**

1. Reviews Power Query M code
2. Applies query folding optimizations
3. Adds Table.Buffer() where appropriate
4. Updates VBA with try-catch patterns
5. Adds error logging to worksheet
6. Exports updated code to version control

### Example 2: Migration and Consolidation

**User:** "I have 5 Excel files with similar Power Query code. Consolidate them into a master workbook with shared sources."

**Copilot Actions:**

1. Exports all queries from 5 files to temp directories
2. Analyzes queries to identify common patterns
3. Creates shared source queries with buffering
4. Refactors individual queries to reference shared sources
5. Imports all queries to master workbook in dependency order
6. Creates documentation of query relationships

### Example 3: Code Quality Improvement

**User:** "Review all Power Query and VBA code in analysis.xlsm and apply best practices."

**Copilot Actions:**

1. Lists all queries and modules
2. Reviews each for performance, security, maintainability
3. Suggests improvements with explanations
4. Applies fixes with user approval
5. Generates before/after comparison
6. Updates version control files

## Resources and Documentation

### Official Documentation

- **Excel MCP Server GitHub:** https://github.com/sbroenne/mcp-server-excel
- **MCP Protocol:** https://modelcontextprotocol.io
- **Power Query M Reference:** https://learn.microsoft.com/power-query/
- **VBA Reference:** https://learn.microsoft.com/office/vba/api/overview/excel

### Community Resources

- **Power Query Community:** Microsoft Tech Community forums
- **Excel VBA Forums:** Stack Overflow, r/excel
- **MCP Community:** MCP Discord server

### Learning Path

1. **Start with Basic Operations:** File creation, worksheet management, named ranges
2. **Progress to Power Query:** Learn M language basics, transformations, query folding
3. **Add VBA Skills:** Error handling, automation, user interaction
4. **Master Advanced Patterns:** Incremental loading, multi-source integration, optimization
5. **Integrate with Tools:** Git workflows, Power BI pipelines, Azure connectivity

## Conclusion

The Excel MCP Server transforms Excel development into a conversational, AI-assisted workflow. By combining GitHub Copilot's natural language understanding with Excel's powerful capabilities, you can:

- **Develop faster** with AI-generated code and transformations
- **Maintain quality** through automated code reviews and optimization
- **Version control** Excel code like any software project
- **Collaborate effectively** with text-based Power Query and VBA files
- **Scale solutions** with reusable patterns and shared components

Start with simple tasks, build confidence, and progressively tackle more complex automation scenarios. The AI assistant is your pair programmer for Excel development.

---

**Quick Start Checklist:**

- [ ] Install .NET SDK 10.0+
- [ ] Configure `.vscode/mcp.json` with Excel MCP Server
- [ ] Verify MCP tools available in Copilot Chat
- [ ] Close all Excel processes before operations
- [ ] Create your first query with AI assistance
- [ ] Export code to version control
- [ ] Review and optimize with Copilot
- [ ] Build your first automated pipeline

Happy Excel Development! 🚀

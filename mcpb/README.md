# Excel MCP Server

**Automate Microsoft Excel with Claude** - Control the installed Excel desktop
application through natural language conversations on Windows x64 or Apple
Silicon macOS.

> **macOS support is experimental beta.** Power Query, VBA, Data Model/DAX/OLAP,
> Tables, PivotTables, charts, slicers, connections, QueryTables, XML Maps,
> screenshots, advanced visual formatting, and Python result reads are not
> supported. See [macOS beta limitations](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta);
> test on copies of important workbooks. Windows retains the complete backend.

Choose the MCPB matching your computer:

- `excel-mcp-<version>.mcpb` (Windows; launches the npm `@latest` package)
- `excel-mcp-<version>-macos-arm64.mcpb`

## What It Does

Excel MCP Server lets you automate Excel through conversation with Claude.
The following complete feature list describes Windows:

- **Create & Edit** - Build spreadsheets, tables, and formulas
- **Analyze Data** - PivotTables, charts, and DAX calculations
- **Transform Data** - Power Query imports and transformations
- **Format & Style** - Conditional formatting, number formats, table styles
- **Automate** - VBA macros, batch operations, data refresh
- **Agent Mode** - Say "show me Excel" and watch AI work in real-time, side-by-side with Claude

The MCP Server provides **31 tools with 326 operations**. Windows exposes the
complete operation set. Apple Silicon macOS exposes only actions marked enabled
in the [generated capability inventory](../docs/MACOS-ACTION-INVENTORY.md).
That enabled set includes workbook lifecycle, worksheet management, values and
formulas, number formats, row and column sizing, merged cells, cell locking,
named ranges, Goal Seek, Data Tables, calculation, and Python formula writes.
Other actions return an explicit unsupported-platform error.

## Requirements

- **Windows x64** with Microsoft Excel 2016 or later, or
- **Apple Silicon macOS** with Excel for Mac 16.112 or later
- **Claude Desktop**
- **Node.js 18+ with npm/npx** for the Windows bundle

## Installation

1. Download the `.mcpb` matching your platform from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Double-click to install in Claude Desktop
3. Restart Claude Desktop if prompted

That's it! Start a new conversation and ask Claude to work with Excel.

## Usage Examples

These Windows examples require the described input data, CSV, or existing
objects. The first creates a new workbook. On Mac, start with plain range
writes/formulas and an absolute native path instead of Tables or analytics.

### Example 1: Create a Sales Tracker (Windows)

**You say:** *"Create a new Excel file called SalesTracker.xlsx with a table for tracking sales. Include columns for Date, Product, Quantity, Unit Price, and Total. Add some sample data and a formula for the Total column."*

**What happens:**
- Creates a new workbook
- Adds column headers (Date, Product, Quantity, Unit Price, Total)
- Enters sample sales data
- Creates formulas in the Total column (Quantity × Unit Price)
- Formats the data as an Excel Table
- Confirms completion with file location

### Example 2: Build a Dashboard with PivotTable and Chart (Windows)

**You say:** *"I want to analyze this data. Create a PivotTable that shows total sales by Product, then add a bar chart to visualize the results."*

**What happens:**
- Creates a PivotTable from the data
- Configures Product as rows and Total as sum values
- Creates a new worksheet for the PivotTable
- Adds a bar chart based on the PivotTable
- Returns confirmation with locations of both

### Example 3: Power Query and Data Model Analysis (Windows)

**You say:** *"Use Power Query to import this CSV file: C:\Data\products.csv. Add the data to the Data Model and create measures for Total Revenue and Average Rating."*

**What happens:**
- Imports the CSV using Power Query
- Loads the data to a worksheet as an Excel Table
- Adds the table to the Power Pivot Data Model
- Creates DAX measures for analysis
- Confirms the data is ready for PivotTable analysis

---

**More Windows examples** (plain cell writes also work on Mac):

- *"Show me Excel side-by-side while you build this dashboard"* - Agent Mode: watch every step happen live
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Create a relationship between the Orders and Products tables using ProductID"*
- *"Run the UpdatePrices macro"*
- *"Show me Excel while you work"* - watch changes in real-time

## Tips for Best Results

- **Be specific** - Include file paths, sheet names, and column references when you know them
- **Start simple** - Build complex spreadsheets step by step
- **Ask to see Excel on Windows** - Say *"Show me Excel while you work"*;
  window/Agent Mode actions are unavailable in the Mac beta
- **Select the exact workbook** - Reuse its existing session when possible.
  Reconcile conflicts for that file only; never close unrelated workbooks or
  terminate shared Mac Excel

## Privacy & Security

Excel MCP Server drives the Excel application on your computer. Workbook files
stay on your local filesystem, but requested tool results are returned to
Claude through the MCP client.

**Optional network features:** Remote M/DAX formatting sends only the supplied
code and requires explicit consent. Python in Excel runs Python code and
referenced worksheet data in Microsoft's cloud.

**Anonymous telemetry:** The MCP Server collects tool usage, performance, and
error-rate metrics. Telemetry excludes file contents, file names, paths, and
personal data.

See our complete [Privacy Policy](https://excelmcpserver.dev/privacy/).

## Troubleshooting

**Claude says the tool isn't available:**
- Restart Claude Desktop after installation
- Check Settings → Integrations to verify Excel MCP Server is enabled

**Excel operations fail:**
- Inspect the matching session and reported error before retrying; reconcile
  only the intended workbook
- Ensure Excel is installed and working normally
- On macOS, verify Excel Automation permission is already granted; ExcelMcp
  never clicks permission prompts or weakens security settings
- For Mac `RecoveryRequired`, reconcile the exact workbook and dialogs manually;
  an uncertain open is not safe to repeat or clean up automatically

**Need help?**
- [Report an issue](https://github.com/sbroenne/mcp-server-excel/issues)
- [Full documentation](https://excelmcpserver.dev/)

## Links

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Feature Reference](https://excelmcpserver.dev/features/)
- [Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/README.md) - Cross-platform AI guidance
- [Privacy Policy](https://excelmcpserver.dev/privacy/)
- [License (MIT)](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

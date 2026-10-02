# Excel (Windows)

**Automate Microsoft Excel with Claude** - Control Excel through natural language conversations. Requires Windows and local Office install.

## What It Does

Excel MCP Server lets you automate Excel through conversation with Claude:

- **Create & Edit** - Build spreadsheets, tables, and formulas
- **Analyze Data** - PivotTables, charts, and DAX calculations
- **Transform Data** - Power Query imports and transformations
- **Format & Style** - Conditional formatting, number formats, table styles
- **Automate** - VBA macros, batch operations, data refresh
- **Agent Mode** - Say "show me Excel" and watch AI work in real-time, side-by-side with Claude

**31 tools with 326 operations** for comprehensive Excel automation.

## Requirements

- **Windows** (required - uses Excel COM automation)
- **Microsoft Excel 2016 or later**
- **Claude Desktop** (Windows version)
- **Node.js 18+ with npm/npx on PATH** (install the current Node.js LTS)
- **An interactive desktop** and network access for package downloads

## Installation

1. Install [Node.js LTS](https://nodejs.org/) if `npx` is not already available.
   Restart Claude Desktop after changing PATH.
2. Download the `.mcpb` file from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
3. Double-click to install in Claude Desktop
4. Restart Claude Desktop if prompted

That's it! Start a new conversation and ask Claude to work with Excel.

## Updates

The bundle tells Claude to run
`npx -y @sbroenne/mcp-server-excel@latest` directly. It contains no custom
launcher, bundled npm, or fixed server executable. Claude's built-in Node.js
does not guarantee availability of the external `npx` command.

On each new server launch, npm resolves the `latest` tag using normal caching
and configuration. Network access is needed for downloads and update checks;
this is not a guaranteed fresh online check every time. A running server is
not replaced automatically.

Before restarting, finish work and explicitly save and close the intended
workbook sessions. Older binary MCPB installations need a one-time installation
of this npx-based bundle; restarting an old bundle does not migrate it.
Changes to bundle metadata/configuration still require manually installing
a new `.mcpb`. To uninstall, remove Excel from Claude's Settings > Extensions.

## Usage Examples

These examples work with any Excel file, including a new empty workbook.

### Example 1: Create a Sales Tracker

**You say:** *"Create a new Excel file called SalesTracker.xlsx with a table for tracking sales. Include columns for Date, Product, Quantity, Unit Price, and Total. Add some sample data and a formula for the Total column."*

**What happens:**
- Creates a new workbook
- Adds column headers (Date, Product, Quantity, Unit Price, Total)
- Enters sample sales data
- Creates formulas in the Total column (Quantity × Unit Price)
- Formats the data as an Excel Table
- Confirms completion with file location

### Example 2: Build a Dashboard with PivotTable and Chart

**You say:** *"I want to analyze this data. Create a PivotTable that shows total sales by Product, then add a bar chart to visualize the results."*

**What happens:**
- Creates a PivotTable from the data
- Configures Product as rows and Total as sum values
- Creates a new worksheet for the PivotTable
- Adds a bar chart based on the PivotTable
- Returns confirmation with locations of both

### Example 3: Power Query and Data Model Analysis

**You say:** *"Use Power Query to import this CSV file: C:\Data\products.csv. Add the data to the Data Model and create measures for Total Revenue and Average Rating."*

**What happens:**
- Imports the CSV using Power Query
- Loads the data to a worksheet as an Excel Table
- Adds the table to the Power Pivot Data Model
- Creates DAX measures for analysis
- Confirms the data is ready for PivotTable analysis

---

**More things you can ask:**

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
- **Ask to see Excel** - Say *"Show me Excel while you work"* to watch changes in real-time
- **Close files first** - Excel MCP needs exclusive access to workbooks during automation

## Privacy & Security

Excel MCP Server drives the Excel application on your computer. Workbook files
stay on your local filesystem, but requested tool results are returned to
Claude through the MCP client.

**Optional network features:** Remote M/DAX formatting sends only the supplied
code and requires explicit consent. Python in Excel runs Python code and
referenced worksheet data in Microsoft's cloud.

**Package downloads:** npx contacts the npm registry to resolve and download the
server. This does not upload workbook contents to npm.

**Anonymous telemetry:** The MCP Server collects tool usage, performance, and
error-rate metrics. Telemetry excludes file contents, file names, paths, and
personal data.

See our complete [Privacy Policy](https://excelmcpserver.dev/privacy/).

## Troubleshooting

**Claude says the tool isn't available:**
- Restart Claude Desktop after installation
- Check Settings → Extensions to verify Excel MCP Server is enabled
- Run `npx -y @sbroenne/mcp-server-excel@latest --version` in PowerShell.
  If npx is missing, install Node.js LTS and restart Claude Desktop.

**Excel operations fail:**
- Close the workbook in Excel before asking Claude to modify it
- Ensure Excel is installed and working normally

**Need help?**
- [Report an issue](https://github.com/sbroenne/mcp-server-excel/issues)
- [Full documentation](https://excelmcpserver.dev/)

## Links

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Feature Reference](https://excelmcpserver.dev/features/)
- [Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/AGENT-SKILLS.md) - Cross-platform AI guidance
- [Privacy Policy](https://excelmcpserver.dev/privacy/)
- [License (MIT)](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

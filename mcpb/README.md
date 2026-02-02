# Excel (Windows)

**Automate Microsoft Excel with Claude** - Control Excel through natural language conversations. Requires Windows and local Office install.

## What It Does

Excel MCP Server lets you automate Excel through conversation with Claude:

- **Create & Edit** - Build spreadsheets, tables, and formulas
- **Analyze Data** - PivotTables, charts, and DAX calculations
- **Transform Data** - Power Query imports and transformations
- **Format & Style** - Conditional formatting, number formats, table styles
- **Automate** - VBA macros, batch operations, data refresh

**22 tools with 210 operations** for comprehensive Excel automation.

## Requirements

- **Windows** (required - uses Excel COM automation)
- **Microsoft Excel 2016 or later**
- **Claude Desktop** (Windows version)

## Installation

1. Download the `.mcpb` file from the [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Double-click to install in Claude Desktop
3. Restart Claude Desktop if prompted

That's it! Start a new conversation and ask Claude to work with Excel.

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

**You say:** *"Use Power Query to import this CSV file: C:/Data/products.csv. Add the data to the Data Model and create measures for Total Revenue and Average Rating."*

**What happens:**
- Imports the CSV using Power Query
- Loads the data to a worksheet as an Excel Table
- Adds the table to the Power Pivot Data Model
- Creates DAX measures for analysis
- Confirms the data is ready for PivotTable analysis

---

**More things you can ask:**

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

Excel MCP Server runs **entirely on your computer**. Your Excel data:
- Never leaves your machine
- Is not sent to any external servers
- Is not used for training AI models

**Anonymous Telemetry:** We collect anonymous usage statistics (tool usage, performance metrics, error rates) to improve the software. No file contents, file names, or personal data are collected.

See our complete [Privacy Policy](https://excelmcpserver.dev/privacy/).

## Troubleshooting

**Claude says the tool isn't available:**
- Restart Claude Desktop after installation
- Check Settings → Integrations to verify Excel MCP Server is enabled

**Excel operations fail:**
- Close the workbook in Excel before asking Claude to modify it
- Ensure Excel is installed and working normally

**Need help?**
- [Report an issue](https://github.com/sbroenne/mcp-server-excel/issues)
- [Full documentation](https://excelmcpserver.dev/)

## Links

- [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
- [Feature Reference](https://excelmcpserver.dev/features/)
- [Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/README.md) - Cross-platform AI guidance
- [Privacy Policy](https://excelmcpserver.dev/privacy/)
- [License (MIT)](https://github.com/sbroenne/mcp-server-excel/blob/main/LICENSE)

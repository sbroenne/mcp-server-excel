# Excel Automation Examples & Use Cases

Excel MCP Server lets AI assistants and coding agents automate the real Microsoft
Excel application using natural-language requests.

**Platform scope:** Windows supports all workflows below. Apple Silicon macOS
is an **experimental beta**: basic cell/formula/number-format work, worksheet
lifecycle, named ranges, sizing, Goal Seek, Data Tables, and licensed Python
formula writes are enabled. Tables, PivotTables/charts, Power Query/DAX, VBA,
advanced visual formatting, Scenarios, and window/Agent Mode operations are
unavailable. See [macOS support and limitations](../specs/MACOS-SUPPORT.md).

## Example prompts

### Create and populate data

- *"Create a new Excel file called SalesTracker.xlsx with a table for Date,
  Product, Quantity, Unit Price, and Total, including sample data."*
- *"Put this data in A1:C4: Name, Age, City / Alice, 30, Seattle / Bob, 25,
  Portland."*
- *"Add a formula column that calculates Quantity times Unit Price."*

On Mac, use plain worksheet cells rather than requesting an Excel Table.

### Analyze and visualize

- *"Create a PivotTable from this data showing total sales by Product, then add a
  bar chart."*
- *"Use Goal Seek to find the price that makes profit equal $100,000, then save
  optimistic and conservative scenarios."*
- *"Create a two-variable data table showing profit for different prices and sales
  volumes."*
- *"Use Power Query to import products.csv, load it to the Data Model, and create
  a measure for Total Revenue."*
- *"Create a slicer for the Region field so I can filter the PivotTable
  interactively."*
- *"Create a relationship between the Orders and Products tables using
  ProductID."*

Goal Seek and Data Tables work in the Mac beta, but saving scenarios does not.
The PivotTable/chart, Power Query, slicer, and relationship prompts require Windows.

### Format and style

- *"Format the Price column as currency and highlight values over $500 in green."*
- *"Convert this range to an Excel Table with a blue style and add a totals row."*
- *"Make the headers bold with a dark background and auto-fit column widths."*
- *"Apply the same section-header styling to A1:G1, A12:G12, and A24:G24 in one
  step."*

Number display formats use the `range` tool. Visual styling, validation, sizing,
and auto-fit use `range_format`.
On Mac, currency number formats and auto-fit are enabled; conditional
highlighting, rich header styling, Table styles, and `format-ranges` are not.

### Automate with code

- *"Export all Power Query M code to files for version control."*
- *"Run the UpdatePrices macro."*
- *"Write a Python in Excel formula that uses pandas to summarize this table."*

## Watch the agent work

Excel normally runs hidden for faster automation. Ask the agent to make it
visible whenever you want to inspect progress:

- *"Show me Excel while you work."*
- *"Show me Excel side-by-side while you build this dashboard."*
- *"Let me watch while you create the chart."*

ExcelMcp can arrange Excel beside the AI assistant and display live progress in
Excel's status bar.
This hidden-window/side-by-side/status-bar experience is **Windows-only**.
Mac uses shared desktop Excel; window and Agent Mode actions remain gated.

## Who should use ExcelMcp?

ExcelMcp is designed for:

- **Data analysts** automating repetitive Excel workflows
- **Developers** building Excel-based data solutions
- **Business users** managing complex workbooks
- **Teams** maintaining Power Query, VBA, and DAX code in version control

It is not designed for:

- Linux environments or Intel Macs
- macOS workflows that require the Windows-only Power Query, Data Model, VBA,
  PivotTable, chart, conditional-formatting, or window-management operations
- Server-side processing without an interactive desktop and Microsoft Excel
- High-volume, Excel-free batch processing where libraries such as ClosedXML or
  EPPlus are a better fit

## Explore the capabilities

- [Data & Analytics](features/DATA-ANALYTICS.md)
- [Cells & Workbooks](features/CELLS-WORKBOOKS.md)
- [Charts & Visualization](features/CHARTS-VISUALS.md)
- [Automation & Advanced](features/AUTOMATION-ADVANCED.md)

[Install ExcelMcp](INSTALLATION.md) when you are ready to try these workflows.

# Excel Automation Examples & Use Cases

Excel MCP Server lets AI assistants and coding agents automate the real Microsoft
Excel application using natural-language requests.

## Example prompts

### Create and populate data

- *"Create a new Excel file called SalesTracker.xlsx with a table for Date,
  Product, Quantity, Unit Price, and Total, including sample data."*
- *"Put this data in A1:C4: Name, Age, City / Alice, 30, Seattle / Bob, 25,
  Portland."*
- *"Add a formula column that calculates Quantity times Unit Price."*

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

### Format and style

- *"Format the Price column as currency and highlight values over $500 in green."*
- *"Convert this range to an Excel Table with a blue style and add a totals row."*
- *"Make the headers bold with a dark background and auto-fit column widths."*
- *"Apply the same section-header styling to A1:G1, A12:G12, and A24:G24 in one
  step."*

Number display formats use the `range` tool. Visual styling, validation, sizing,
and auto-fit use `range_format`.

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

## Who should use ExcelMcp?

ExcelMcp is designed for:

- **Data analysts** automating repetitive Excel workflows
- **Developers** building Excel-based data solutions
- **Business users** managing complex workbooks
- **Teams** maintaining Power Query, VBA, and DAX code in version control

It is not designed for:

- Linux or macOS environments
- Server-side processing without an interactive desktop and Microsoft Excel
- High-volume, Excel-free batch processing where libraries such as ClosedXML or
  EPPlus are a better fit

## Explore the capabilities

- [Data & Analytics](features/DATA-ANALYTICS.md)
- [Cells & Workbooks](features/CELLS-WORKBOOKS.md)
- [Charts & Visualization](features/CHARTS-VISUALS.md)
- [Automation & Advanced](features/AUTOMATION-ADVANCED.md)

[Install ExcelMcp](INSTALLATION.md) when you are ready to try these workflows.

# ExcelMcp - What You Can Automate

**31 feature areas with 394 operations, exposed through 60 MCP tools and the CLI**

ExcelMcp uses the installed Microsoft Excel application, not a file parser.
Excel itself calculates formulas, refreshes data, runs macros, and renders
charts. The MCP Server and CLI provide the same capabilities.

## Explore by goal

| What you want to accomplish | Feature area |
|---|---|
| Import and transform data, connect tables, and build analytical summaries | [Data & Analytics](docs/features/DATA-ANALYTICS.md) |
| Update cells and formulas, format reports, and manage sheets and files | [Cells & Workbooks](docs/features/CELLS-WORKBOOKS.md) |
| Create charts, interactive filters, and other worksheet visuals | [Charts & Visualization](docs/features/CHARTS-VISUALS.md) |
| Run code, compare assumptions, control Excel windows, and exchange XML data | [Automation & Advanced](docs/features/AUTOMATION-ADVANCED.md) |

These pages describe capabilities and important limitations, not command
syntax. Ask your AI assistant for the result you want in plain language.
For current actions and inputs, use the MCP tool descriptions or the CLI's
built-in help: `excelcli --help`, then `excelcli <command> --help`.

## Practical guides

- [Refresh Power Query from an AI assistant](docs/guides/REFRESH-POWER-QUERY.md)
- [Build and update PivotTables with an AI assistant](docs/guides/AUTOMATE-PIVOTTABLES.md)
- [Query the Excel Data Model with DAX](docs/guides/QUERY-DATA-MODEL-WITH-DAX.md)
- [Run VBA macros from an AI agent](docs/guides/RUN-VBA-MACROS.md)
- [Real Excel automation vs. file-parser libraries](docs/guides/EXCEL-COM-VS-FILE-PARSERS.md)

The [workflow guidance](docs/reference/README.md) explains sequencing, safe
edits, recovery, and decisions that command help alone cannot teach.

## Requirements and boundaries

ExcelMcp requires Windows and installed desktop Excel. Some capabilities also
depend on the Excel version, account licensing, data providers, or an interactive
desktop; the category pages call out these requirements.

Edits affect the live workbook. Saving, discarding changes, refreshing external
sources, and running code have different consequences. See
[working safely with Excel](docs/reference/behavioral-rules.md) before planning
an unattended workflow.

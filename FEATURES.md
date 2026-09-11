# ExcelMcp - Complete Feature Reference

**31 specialized tools with 326 operations for comprehensive Excel automation**

Excel MCP Server automates the real Microsoft Excel application through four focused capability areas. Start with the category that matches your goal, or use the quick reference below to find a tool for a specific task.

## Explore by goal

| Goal | Feature area | Included tools |
|---|---|---|
| Import, transform, model, and summarize data | [Data & Analytics](docs/features/DATA-ANALYTICS.md) | Power Query, Data Model & DAX, Excel Tables, PivotTables, Data Connections, QueryTables |
| Read, write, and format cells; manage formulas, sheets, and files | [Cells & Workbooks](docs/features/CELLS-WORKBOOKS.md) | File Operations, Calculation, Ranges, Worksheets, Workbook, Named Ranges |
| Build charts and interactive visuals; capture workbook screenshots | [Charts & Visualization](docs/features/CHARTS-VISUALS.md) | Charts, Slicers, Conditional Formatting, Screenshots, Drawing Objects, Sparklines |
| Run VBA or Python, control Excel windows, solve What-If scenarios, and work with XML Maps | [Automation & Advanced](docs/features/AUTOMATION-ADVANCED.md) | VBA, Python in Excel, Window Management, What-If Analysis, XML Maps |

> **New to Excel MCP Server?** You do not need to memorize operation names. Describe the result you want in plain language and your AI assistant selects the appropriate tool.

## Task guides

Prefer a walkthrough to a reference table? The [task guides](docs/guides/README.md)
cover the most common jobs end to end:

- [Refresh Power Query from an AI assistant](docs/guides/REFRESH-POWER-QUERY.md)
- [Build and update PivotTables with an AI assistant](docs/guides/AUTOMATE-PIVOTTABLES.md)
- [Query the Excel Data Model with DAX](docs/guides/QUERY-DATA-MODEL-WITH-DAX.md)
- [Run VBA macros from an AI agent](docs/guides/RUN-VBA-MACROS.md)
- [Real Excel automation vs. file-parser libraries](docs/guides/EXCEL-COM-VS-FILE-PARSERS.md)

---

## 🔧 Tool Selection Quick Reference

| Task | Tool | Feature reference |
|------|------|-------------------|
| Import or transform data | `powerquery`; `connection` for existing OLEDB/ODBC sources; `querytable` for direct text/web imports | [Data & Analytics](docs/features/DATA-ANALYTICS.md) |
| Build a Power Pivot model and DAX measures | `datamodel` | [Data Model & DAX](https://excelmcpserver.dev/features/data-analytics/#data-model-dax-power-pivot) |
| Create or update a PivotTable for aggregation | `pivottable` | [PivotTables](https://excelmcpserver.dev/features/data-analytics/#pivottables) |
| Find an input for a target result, compare scenarios, or build What-If data tables | `analysis` | [What-If Analysis](https://excelmcpserver.dev/features/automation-advanced/#what-if-analysis) |
| Visualize data | `chart` | [Charts](https://excelmcpserver.dev/features/charts-visuals/#charts) |
| Update parameters | `namedrange` (write operation) | [Cells & Workbooks](docs/features/CELLS-WORKBOOKS.md) |
| Read or write cells and manage formulas | `range` (including `set-formulas`) | [Ranges](https://excelmcpserver.dev/features/cells-workbooks/#ranges) |
| Format or validate data | `range_format` (`format-range`, `format-ranges`, `validate-range`) | [Ranges](https://excelmcpserver.dev/features/cells-workbooks/#ranges) |
| Run a macro | `vba` | [VBA Macros](https://excelmcpserver.dev/features/automation-advanced/#vba-macros) |

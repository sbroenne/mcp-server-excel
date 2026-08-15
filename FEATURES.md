# ExcelMcp - Complete Feature Reference

**31 specialized tools with 326 operations for comprehensive Excel automation**

Excel MCP Server automates the real Microsoft Excel application through four focused capability areas. Start with the category that matches your goal, or use the quick reference below.

## Explore by goal

| Goal | Feature area | Included tools |
|---|---|---|
| Import, transform, model, and summarize data | [Data & Analytics](docs/features/DATA-ANALYTICS.md) | Power Query, Data Model & DAX, Excel Tables, PivotTables, Data Connections, QueryTables |
| Edit cells, formulas, sheets, and files | [Cells & Workbooks](docs/features/CELLS-WORKBOOKS.md) | File Operations, Calculation, Ranges, Worksheets, Workbook, Named Ranges |
| Build presentation-ready workbook visuals | [Charts & Visualization](docs/features/CHARTS-VISUALS.md) | Charts, Slicers, Conditional Formatting, Screenshots, Drawing Objects, Sparklines |
| Run code and specialized Excel workflows | [Automation & Advanced](docs/features/AUTOMATION-ADVANCED.md) | VBA, Python in Excel, Window Management, What-If Analysis, XML Maps |

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

| Task | Tool |
|------|------|
| Import data | `powerquery` or `connection` |
| Create analysis | `analysis` for Goal Seek/scenarios/data tables; `pivottable` for aggregation |
| Visualize data | `chart` |
| Update parameters | `namedrange` (write operation) |
| Manage formulas | `range` (set-formulas) |
| Format data | `range` / `range_format` (`format-range`, `format-ranges`, `validate-range`) |
| Script automation | `vba` (run macro) |

---
name: excel-mcp-report-formatting
description: >-
  Apply optional presentation conventions when the user requests Excel report
  formatting or a new reader-facing report through Excel MCP tools. Use for
  readable headers, meaningful number formats, financial-model colours, and
  report layout. Do not load for ordinary reads, raw exports, targeted data
  updates, Power Query or Data Model work, or failure recovery without a
  report-formatting request.
compatibility: Windows with desktop Excel and the Excel MCP Server.
---

# Report formatting with Excel MCP

Use the user's requested style first, the existing template second, and the
optional conventions in the [formatting guide](references/report-formatting.md)
only when neither specifies a style. Read that guide for the requested report,
not unrelated workbook work.

Discover exact actions and parameters from the MCP tool schemas.
Preserve values, formulas, units, identifiers, existing layouts, and unrelated
sheets. Formatting is not permission to refresh data or rebuild workbook objects.
Use `table` styles for Tables, `chart_config` for charts, and `pivottable_field`
for PivotTable number formats; do not apply `range_format` to Table or PivotTable
cells instead of their owning controls.

Inspect the affected formats and calculated outputs before saving the authorized
result. Financial-model colours and dashboard visuals are optional, not a
requirement for every report. Screenshots are optional and need a desktop.

General workflow and recovery documentation is available on the
[website](https://excelmcpserver.dev/reference/); it is not another required skill.

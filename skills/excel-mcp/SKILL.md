---
name: excel-mcp
description: Automate Microsoft Excel on Windows via the Excel MCP Server (Power Query, Data Model/DAX, tables, ranges, charts, formatting, VBA).
license: MIT
version: 1.1.5
tags:
  - excel
  - automation
  - mcp
  - windows
  - powerquery
  - dax
  - pivottable
  - vba
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://sbroenne.github.io/mcp-server-excel/
---

# Excel MCP Server Skill

Use this skill when the user needs to create, read, or modify Excel workbooks through the Excel MCP Server on Windows.

## Preconditions
- Windows host with Microsoft Excel installed (2016+).
- Use full Windows paths (for example, `C:\Users\Name\Documents\Report.xlsx`). Do not invent user names or paths.
- Excel files must not be open in another Excel instance.

## Session and Batch Workflow
1. `excel_file` open/create-empty to get a `sessionId`.
2. For 3+ operations, start `excel_batch` to get a `batchId`.
3. Perform operations using `sessionId` or `batchId`.
4. Save and close with `excel_batch` commit or `excel_file` close (save=true).

## Tool Map
- Session and files: `excel_file`, `excel_batch`
- Worksheets: `excel_worksheet`, `excel_worksheet_style`
- Ranges: `excel_range`, `excel_range_edit`, `excel_range_format`, `excel_range_link`
- Tables: `excel_table`, `excel_table_column`
- Power Query: `excel_powerquery`
- Data Model/DAX: `excel_datamodel`, `excel_datamodel_rel`
- PivotTables: `excel_pivottable`, `excel_pivottable_field`, `excel_pivottable_calc`
- Charts: `excel_chart`, `excel_chart_config`
- Formatting and filters: `excel_conditionalformat`, `excel_slicer`
- Connections: `excel_connection`
- Named ranges: `excel_namedrange`
- VBA macros: `excel_vba` (requires VBA trust)

## Core Rules
- Use 2D arrays for values and formulas. A single cell is still `[[value]]`.
- Prefer targeted updates. Avoid delete-and-rebuild workflows.
- List before delete or rename to confirm names.
- Use US number format codes (for example, `#,##0.00`).
- Check `success` and `errorMessage`. Follow `suggestedNextActions` when provided.
- Avoid confirmation loops. Ask only when the request is ambiguous or destructive.

## Data Model Workflow
- Tables must be in the Data Model before DAX measures work.
  - Worksheet tables: `excel_table` with `add-to-datamodel`.
  - External data: `excel_powerquery` with `loadDestination='data-model'`, then `refresh`.
- Use `excel_datamodel` for measures and metadata, `excel_datamodel_rel` for relationships.

## Power Query Workflow
- `create` defines the query; call `refresh` to load data.
- Choose `loadDestination` based on the goal: `worksheet`, `data-model`, `both`, `connection-only`.

## Output Expectations
- Summarize what changed and where.
- Return file paths and session/batch IDs when created.

## Reference Documentation
- references/behavioral-rules.md
- references/anti-patterns.md
- references/claude-desktop.md
- references/excel_powerquery.md
- references/excel_datamodel.md
- references/excel_table.md
- references/excel_range.md
- references/excel_worksheet.md

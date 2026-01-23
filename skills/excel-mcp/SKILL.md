---
name: excel-mcp
description: >
  Automate Microsoft Excel on Windows via COM interop. Use when creating, reading, 
  or modifying Excel workbooks. Supports Power Query (M code), Data Model (DAX measures), 
  PivotTables, Tables, Ranges, Charts, Formatting, VBA macros, and connections.
  Triggers: Excel, spreadsheet, workbook, xlsx, Power Query, DAX, PivotTable, VBA.
license: MIT
version: 1.2.0
tags:
  - excel
  - automation
  - mcp
  - windows
  - powerquery
  - dax
  - pivottable
  - vba
  - data-model
  - charts
  - formatting
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel MCP Server Skill

Server-specific guidance for Excel MCP Server. Tools are auto-discovered - this documents quirks, workflows, and gotchas.

## Preconditions

- Windows host with Microsoft Excel installed (2016+)
- Use full Windows paths: `C:\Users\Name\Documents\Report.xlsx`
- Excel files must not be open in another Excel instance
- VBA operations require "Trust access to VBA project object model" enabled in Excel Trust Center

## Session Workflow

1. **Open/Create**: `excel_file(open)` or `excel_file(create-empty)` → returns `sessionId`
2. **Perform Operations**: Pass `sessionId` to all tool calls
3. **Save and Close**: `excel_file(close, save=true)` to persist changes

**Session Tips:**
- Call `excel_file(list)` first to check for existing sessions (reuse if file already open)
- Check `canClose=true` before closing (no active operations running)
- Without `save=true`, all changes are discarded

## Format Cells After Setting Values (CRITICAL)

Without formatting, dates appear as serial numbers and currency as plain numbers.

| Data Type | Format Code | Example Result |
|-----------|-------------|----------------|
| Currency (USD) | `$#,##0.00` | $1,234.56 |
| Currency (EUR) | `€#,##0.00` | €1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |
| Number | `#,##0.00` | 1,234.56 |

**Workflow**: `excel_range(set-values)` → `excel_range_format(set-number-format)`

## Format Data as Excel Tables (CRITICAL)

Always convert tabular data to Excel Tables, not plain ranges:

```
1. excel_range(set-values)  # Write data with headers
2. excel_table(create)      # Convert to Table
```

**Benefits**: Structured references, auto-expand, built-in filtering, required for Data Model.

## Core Rules

1. **2D Arrays**: Values and formulas use 2D arrays. Single cell = `[[value]]`
2. **Targeted Updates**: Modify specific cells, not entire structures
3. **List Before Delete**: Verify names exist before delete/rename operations
4. **US Format Codes**: Use US locale format codes (`#,##0.00` not `#.##0,00`)
5. **Check Results**: Always check `success` and `errorMessage` in responses
6. **Follow Suggestions**: Act on `suggestedNextActions` when provided

## Data Model Workflow

Tables must be in the Data Model before DAX measures work:

1. **Worksheet Tables**: `excel_table(add-to-datamodel)`
2. **External Data**: `excel_powerquery(create, loadDestination='data-model')` → `excel_powerquery(refresh)`
3. **Create Measures**: `excel_datamodel(create-measure)`

**Power Pivot Limitations** (vs Power BI):
- NO calculated tables - use Power Query instead
- NO calculated columns via COM API - use Power Query or DAX measures
- Measures and relationships work fully

## Power Query Workflow

1. **Create Query**: `excel_powerquery(create, mCode='...')` - imports M code
2. **Load Data**: `excel_powerquery(refresh, refreshTimeoutSeconds=120)` - REQUIRED parameter
3. **Load Destinations**: `worksheet`, `data-model`, `both`, or `connection-only`

**Server Quirks:**
- `refresh` REQUIRES `refreshTimeoutSeconds` (60-600 seconds) - will fail without it
- M code is auto-formatted on create/update
- `update` action auto-refreshes after updating M code

## Star Schema Design

### Architecture
```
Power Query (ETL):          DAX (Business Logic):
- Load/transform data       - Calculate measures
- Create fact tables        - Time intelligence
- Create dimension tables   - Business rules
```

### Relationships
- Use `excel_datamodel_rel` to create relationships
- Pattern: Fact[ForeignKey] → Dimension[PrimaryKey]

## Named Ranges for Parameters

Use `excel_namedrange` for values Power Query can read:

1. `excel_worksheet(create)` - e.g., "_Setup"
2. `excel_range(set-values)` - parameter values
3. `excel_namedrange(create)` - named reference
4. M code: `Excel.CurrentWorkbook(){[Name = "Param_Name"]}[Content]{0}[Column1]`

## Common Patterns

### Import CSV and Build Dashboard
```
excel_file(open) → sessionId
excel_powerquery(create, loadDestination='data-model', mCode=Csv.Document(...))
excel_powerquery(refresh, refreshTimeoutSeconds=120)
excel_datamodel(create-measure, daxFormula='SUM(...)')
excel_pivottable(create-from-datamodel)
excel_file(close, save=true)
```

### Update Existing Workbook
```
excel_file(open) → sessionId
excel_powerquery(list)  # Check existing
excel_powerquery(update)  # Auto-refreshes
excel_range(get-values)   # Verify
excel_file(close, save=true)
```

## Reference Documentation

See `references/` for detailed guidance:

- @references/workflows.md - Production patterns
- @references/behavioral-rules.md - Execution guidelines
- @references/anti-patterns.md - Common mistakes
- @references/excel_powerquery.md - Power Query specifics
- @references/excel_datamodel.md - Data Model/DAX specifics
- @references/excel_table.md - Table operations
- @references/excel_range.md - Range operations
- @references/excel_worksheet.md - Worksheet operations
- @references/excel_chart.md - Charts, formatting, and trendlines
- @references/claude-desktop.md - Claude Desktop setup

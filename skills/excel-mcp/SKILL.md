---
name: excel-mcp
description: >
  Automate Microsoft Excel on Windows via COM interop. Use when creating, reading,
  or modifying Excel workbooks. Supports Power Query (M code), Data Model (DAX measures),
  PivotTables, Tables, Ranges, Charts, Slicers, Formatting, VBA macros, and connections.
  Triggers: Excel, spreadsheet, workbook, xlsx, Power Query, DAX, PivotTable, VBA.
compatibility: Windows + Microsoft Excel 2016+ required. Uses COM interop - does NOT work on macOS or Linux.
license: MIT
version: 1.3.0
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel MCP Server Skill

Provides 200+ Excel operations via Model Context Protocol. Tools are auto-discovered - this documents quirks, workflows, and gotchas.

## Preconditions

- Windows host with Microsoft Excel installed (2016+)
- Use full Windows paths: `C:\Users\Name\Documents\Report.xlsx`
- Excel files must not be open in another Excel instance

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: NEVER Ask Clarifying Questions

**STOP.** If you're about to ask "Which file?", "What table?", "Where should I put this?" - DON'T.

| Bad (Asking) | Good (Discovering) |
|--------------|-------------------|
| "Which Excel file should I use?" | `excel_file(list)` → use the open session |
| "What's the table name?" | `excel_table(list)` → discover tables |
| "Which sheet has the data?" | `excel_worksheet(list)` → check all sheets |
| "Should I create a PivotTable?" | YES - create it on a new sheet |

**You have tools to answer your own questions. USE THEM.**

### Rule 2: Format Data Professionally

Always apply number formats after setting values:

| Data Type | Format Code | Result |
|-----------|-------------|--------|
| USD | `$#,##0.00` | $1,234.56 |
| EUR | `€#,##0.00` | €1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |

**Workflow:**
```
1. excel_range set-values (data is now in cells)
2. excel_range_format set-number-format (apply format)
```

### Rule 3: Use Excel Tables (Not Plain Ranges)

Always convert tabular data to Excel Tables:

```
1. excel_range set-values (write data including headers)
2. excel_table create tableName="SalesData" rangeAddress="A1:D100"
```

**Why:** Structured references, auto-expand, required for Data Model/DAX.

### Rule 4: Session Lifecycle

```
1. excel_file(action: 'open', excelPath: '...')  → sessionId
2. All operations use sessionId
3. excel_file(action: 'close', save: true)  → saves and closes
```

**Unclosed sessions leave Excel processes running, locking files.**

### Rule 5: Data Model Prerequisites

DAX operations require tables in the Data Model:

```
Step 1: Create table → Table exists
Step 2: excel_table(action: 'add-to-datamodel') → Table in Data Model
Step 3: excel_datamodel(action: 'create-measure') → NOW this works
```

### Rule 6: Power Query Development Lifecycle

**BEST PRACTICE: Test-First Workflow**

```
1. excel_powerquery(action: 'evaluate', mCode: '...') → Test WITHOUT persisting
2. excel_powerquery(action: 'create', ...) → Store validated query
3. excel_powerquery(action: 'refresh', ...) → Load data
```

**Why evaluate first:**
- Catches syntax errors and missing sources BEFORE creating permanent queries
- Better error messages than COM exceptions from create/update
- See actual data preview (columns + sample rows)
- No cleanup needed - like a REPL for M code
- Skip only for trivial literal tables

**Common mistake:** Creating/updating without evaluate → pollutes workbook with broken queries

### Rule 7: Targeted Updates Over Delete-Rebuild

- **Prefer**: `set-values` on specific range (e.g., `A5:C5` for row 5)
- **Avoid**: Deleting and recreating entire structures

**Why:** Preserves formatting, formulas, and references.

### Rule 8: Follow suggestedNextActions

Error responses include actionable hints:
```json
{
  "success": false,
  "errorMessage": "Table 'Sales' not found in Data Model",
  "suggestedNextActions": ["excel_table(action: 'add-to-datamodel', tableName: 'Sales')"]
}
```

## Tool Selection Quick Reference

| Task | Tool | Key Action |
|------|------|------------|
| Create/open/save workbooks | `excel_file` | open, create, close |
| Write/read cell data | `excel_range` | set-values, get-values |
| Format cells | `excel_range_format` | set-number-format |
| Create tables from data | `excel_table` | create |
| Add table to Power Pivot | `excel_table` | add-to-datamodel |
| Create DAX formulas | `excel_datamodel` | create-measure |
| Create PivotTables | `excel_pivottable` | create, create-from-datamodel |
| Filter with slicers | `excel_slicer` | set-slicer-selection |
| Create charts | `excel_chart` | create-from-range |

## Reference Documentation

See `references/` for detailed guidance:

- @references/behavioral-rules.md - Core execution rules and LLM guidelines
- @references/anti-patterns.md - Common mistakes to avoid
- @references/workflows.md - Data Model constraints and patterns
- @references/excel_chart.md - Charts and formatting
- @references/excel_conditionalformat.md - Conditional formatting operations
- @references/excel_datamodel.md - Data Model/DAX specifics
- @references/excel_powerquery.md - Power Query specifics
- @references/excel_range.md - Range operations and number formats
- @references/excel_slicer.md - Slicer operations
- @references/excel_table.md - Table operations
- @references/excel_worksheet.md - Worksheet operations

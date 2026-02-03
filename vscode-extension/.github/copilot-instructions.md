# Excel MCP Server - Quick Reference

> **When user asks about Excel files, spreadsheets, workbooks, or data in .xlsx/.xlsm files - USE the excel_* tools.**

## When to Use Excel MCP

USE these tools when user wants to:
- Read/write Excel data, formulas, or formatting
- Create PivotTables, charts, or tables
- Import data via Power Query
- Run VBA macros
- Any .xlsx or .xlsm file operations

DO NOT use for: CSV files (use standard file tools), Google Sheets, or non-Excel formats.

---

## Prerequisites

- **Windows OS** - Excel COM automation requires Windows
- **Microsoft Excel 2016+** - Must be installed
- **File CLOSED in Excel** - COM requires exclusive access

---

## Tool Selection (Which Tool for Which Task?)

| Task | Use | NOT |
|------|-----|-----|
| Import external data (CSV, SQL, APIs) | `excel_powerquery` | `excel_table` |
| DAX measures / calculated fields | `excel_datamodel` | `excel_range` |
| Worksheet formulas (SUM, VLOOKUP) | `excel_range` | `excel_datamodel` |
| Structured data with filtering | `excel_table` | `excel_range` |
| Interactive summarization | `excel_pivottable` | `excel_table` |
| VBA automation | `excel_vba` | Requires .xlsm |

**Data Model prerequisite**: Before using `excel_datamodel`, data must be loaded with `loadDestination: 'data-model'` or `'both'` via `excel_powerquery`.

---

## Cross-Tool Workflow Patterns

### Import Data -> Analyze -> Visualize
```
1. excel_file(action: 'open')
2. excel_powerquery(action: 'create', loadDestination: 'worksheet')
3. excel_pivottable(action: 'create-from-table')
4. excel_chart(action: 'create-from-pivottable')
5. excel_file(action: 'close', save: true)
```

### Build DAX Analytics
```
1. excel_file(action: 'open')
2. excel_powerquery(action: 'create', loadDestination: 'data-model')
3. excel_datamodel(action: 'create-measure', formula: 'SUM(...)')
4. excel_pivottable(action: 'create-from-datamodel')
5. excel_file(action: 'close', save: true)
```

### Batch Updates (Multiple Items)
```
# Use bulk data operations - set entire ranges at once:
excel_range(action: 'set-values', values: [[1,2,3], [4,5,6]]) # NOT cell-by-cell
excel_range(action: 'set-formulas', formulas: [['=A1', '=B1']]) # Multiple formulas at once
```

---

## Common Cross-Tool Mistakes

| Mistake | Fix |
|---------|-----|
| Using `excel_table` to import CSV | Use `excel_powerquery` (handles encoding, transforms) |
| Using `excel_range` for DAX | Use `excel_datamodel` (DAX != worksheet formulas) |
| Multiple single-item calls | Use bulk actions when available |
| Closing session between operations | Keep session open until workflow complete |
| Working on file open in Excel | Ask user to close file first |

---

## Session Lifecycle Reminder

```
excel_file(action: 'open') -> [all operations with sessionId] -> excel_file(action: 'close')
```

- **DEFAULT: `showExcel: false`** - Use hidden mode for faster background automation
- Only use `showExcel: true` if user explicitly requests to watch changes
- If `showExcel: true` was used, **ask before closing** (user may want to inspect)
- Use `excel_file(action: 'list')` to check session state if uncertain

---

## For Per-Tool Details

Each tool has detailed documentation in its schema description. For server-specific quirks, the MCP server exposes prompt resources you can request.

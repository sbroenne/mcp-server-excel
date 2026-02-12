# Excel MCP Server - Quick Reference

> **When user asks about Excel files, spreadsheets, workbooks, or data in .xlsx/.xlsm files - USE the Excel MCP tools.**

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
| Import external data (CSV, SQL, APIs) | `powerquery` | `table` |
| DAX measures / calculated fields | `datamodel` | `range` |
| Worksheet formulas (SUM, VLOOKUP) | `range` | `datamodel` |
| Structured data with filtering | `table` | `range` |
| Interactive summarization | `pivottable` | `table` |
| VBA automation | `vba` | Requires .xlsm |

**Data Model prerequisite**: Before using `datamodel`, data must be loaded with `loadDestination: 'data-model'` or `'both'` via `powerquery`.

---

## Cross-Tool Workflow Patterns

### Import Data -> Analyze -> Visualize
```
1. file(action: 'open')
2. powerquery(action: 'create', loadDestination: 'worksheet')
3. pivottable(action: 'create-from-table')
4. chart(action: 'create-from-pivottable')
5. file(action: 'close', save: true)
```

### Build DAX Analytics
```
1. file(action: 'open')
2. powerquery(action: 'create', loadDestination: 'data-model')
3. datamodel(action: 'create-measure', formula: 'SUM(...)')
4. pivottable(action: 'create-from-datamodel')
5. file(action: 'close', save: true)
```

### Batch Updates (Multiple Items)
```
# Use bulk data operations - set entire ranges at once:
range(action: 'set-values', values: [[1,2,3], [4,5,6]])  # NOT cell-by-cell
range(action: 'set-formulas', formulas: [['=A1', '=B1']])  # Multiple formulas at once
```

---

## Common Cross-Tool Mistakes

| Mistake | Fix |
|---------|-----|
| Using `table` to import CSV | Use `powerquery` (handles encoding, transforms) |
| Using `range` for DAX | Use `datamodel` (DAX != worksheet formulas) |
| Multiple single-item calls | Use bulk actions when available |
| Closing session between operations | Keep session open until workflow complete |
| Working on file open in Excel | Ask user to close file first |

---

## Session Lifecycle Reminder

```
file(action: 'open') -> [all operations with sessionId] -> file(action: 'close')
```

- **DEFAULT: `showExcel: false`** - Use hidden mode for faster background automation
- Only use `showExcel: true` if user explicitly requests to watch changes
- If `showExcel: true` was used, **ask before closing** (user may want to inspect)
- Use `file(action: 'list')` to check session state if uncertain

---

## For Per-Tool Details

Each tool has detailed documentation in its schema description. For server-specific quirks, the MCP server exposes prompt resources you can request.

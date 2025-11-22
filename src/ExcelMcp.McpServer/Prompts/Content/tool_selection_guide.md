# Tool Selection Guide

## Pre-Requisite
- File must be CLOSED in Excel (exclusive access required by COM)
- Use `excel_file(open)` to start a session, `excel_file(close)` to end

**CRITICAL - Session Lifetime Rules:**
- **KEEP session open** across multiple operations
- **ONLY close** when user explicitly requests it OR all operations are complete
- **ASK user** "Should I close the session now?" if unclear whether more operations are needed
- Closing mid-workflow = lost session, cannot resume


## Quick Reference

| Need | Use | NOT |
|------|-----|-----|
| External data (databases, APIs, CSV) | `excel_powerquery` + `loadDestination` | `excel_table` (data already in Excel) |
| Connection management | `excel_connection` | - |
| DAX measures / Data Model | `excel_datamodel` | `excel_range` (worksheet formulas) |
| Data in worksheets (values, formulas, format) | `excel_range` | - |
| Convert range to table | `excel_table` | - |
| Sheet lifecycle (create, delete, hide, rename) | `excel_worksheet` | - |
| Named ranges (parameters) | `excel_namedrange` (use `create-bulk` for 2+) | - |
| VBA macros (.xlsm only) | `excel_vba` | - |
| PivotTables | `excel_pivottable` | - |

## Common Mistakes

**Don't use `excel_table` for external data**
- WRONG: `excel_table(create)` for CSV import
- CORRECT: `excel_powerquery(create, loadDestination='worksheet')`

**loadDestination matters**
- WRONG: `excel_powerquery` without `loadDestination` for DAX
- CORRECT: `excel_powerquery(create, loadDestination='data-model')`

**Use bulk operations for multiple items**
- WRONG: `excel_namedrange(create)` called 5 times
- CORRECT: `excel_namedrange(create-bulk)` with JSON array

**DAX is not worksheet formulas**
- WRONG: Using `excel_range` for DAX syntax
- CORRECT: `excel_datamodel(create-measure)` with DAX

**WorksheetAction vs DAX**
- Worksheet formulas: `excel_range` with `=SUM(A1:A10)`
- DAX measures: `excel_datamodel` with `SUM(Sales[Amount])`

# excel_table Tool

**Actions**: list, info, create, rename, delete, resize, set-style, toggle-totals, set-column-total, append, add-to-datamodel, apply-filter, apply-filter-values, clear-filters, get-filters, add-column, remove-column, rename-column, get-structured-reference, sort, sort-multi, get-column-number-format, set-column-number-format

**When to use excel_table**:
- Convert worksheet ranges to Excel Tables (ListObjects)
- Add structure: AutoFilter, structured references ([@Column])
- Data already in Excel worksheets
- Use excel_powerquery for external data sources
- Use excel_datamodel for Power Pivot operations

**Server-specific behavior**:
- Tables have AutoFilter by default
- Structured references: =Table1[@ColumnName]
- add-to-datamodel: Adds table to Power Pivot (requires data model)
- resize: Changes table range boundaries

**Action disambiguation**:
- create: Convert existing range to table
- resize: Expand/shrink table range
- append: Add rows to table
- add-to-datamodel: Make table available for DAX (different from excel_powerquery loadDestination)
- apply-filter: Filter by column criteria
- sort: Single column sort
- sort-multi: Multiple column sort

**Common mistakes**:
- Using excel_table for external data → Use excel_powerquery instead
- Confusing add-to-datamodel with loadDestination → Different concepts
- Not creating table before calling table actions → Use create first

**Workflow optimization**:
- Multiple table operations? Use begin_excel_batch
- Create table → Apply filters → Sort → Format columns
- Use structured references in formulas for maintainability

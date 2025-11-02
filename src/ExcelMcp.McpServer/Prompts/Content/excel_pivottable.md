# excel_pivottable Tool

**Actions**: list, get-info, create-from-range, create-from-table, delete, refresh, list-fields, add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, get-data

**When to use excel_pivottable**:
- Create PivotTables from ranges or tables
- Configure PivotTable fields and calculations
- Use excel_table for source data tables
- Use excel_datamodel for DAX-based analytics

**Server-specific behavior**:
- Requires source data in range or table
- PivotTables auto-refresh when source data changes
- Fields can be rows, columns, values, or filters
- Value field functions: Sum, Count, Average, Max, Min, etc.

**Action disambiguation**:
- create-from-range: Create PivotTable from range address
- create-from-table: Create PivotTable from Excel Table
- add-row-field: Field goes to row area
- add-column-field: Field goes to column area
- add-value-field: Field goes to values area (calculations)
- add-filter-field: Field goes to filter area
- set-field-function: Change aggregation (Sum, Count, etc.)

**Common mistakes**:
- Creating PivotTable without source data → Prepare data first
- Not refreshing after source data changes → Use refresh action
- Wrong field type → Rows vs Columns vs Values have different purposes

**Workflow optimization**:
- Multiple PivotTables? Use begin_excel_batch
- Pattern: Create → Add fields → Set functions → Format → Sort
- Batch mode for configuring multiple fields

# excel_pivottable Tool

**Related tools**:
- excel_range - For source data ranges (use create-from-range)
- excel_table - For structured source tables (use create-from-table)
- excel_datamodel - For Data Model tables with DAX (use create-from-datamodel)
- excel_batch - Use for multiple PivotTable field operations (75-90% faster)

**Actions**: list, get-info, create-from-range, create-from-table, create-from-datamodel, delete, refresh, list-fields, add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, get-data

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
- create-from-range: Create PivotTable from range address (sheetName + range parameters)
- create-from-table: Create PivotTable from Excel Table/ListObject (excelTableName = worksheet table name)
- create-from-datamodel: Create PivotTable from Power Pivot Data Model table (dataModelTableName = Data Model table name)
- add-row-field: Field goes to row area
- add-column-field: Field goes to column area
- add-value-field: Field goes to values area (calculations)
- add-filter-field: Field goes to filter area
- set-field-function: Change aggregation (Sum, Count, etc.)

**Parameter guidance for create actions**:
- create-from-range: sheetName, range, destinationSheet, destinationCell, pivotTableName
- create-from-table: excelTableName (Excel Table name), destinationSheet, destinationCell, pivotTableName
- create-from-datamodel: dataModelTableName (Data Model table name), destinationSheet, destinationCell, pivotTableName

**Common mistakes**:
- Creating PivotTable without source data → Prepare data first
- Not refreshing after source data changes → Use refresh action
- Wrong field type → Rows vs Columns vs Values have different purposes

**Workflow optimization**:
- Multiple PivotTables? Use begin_excel_batch
- Pattern: Create → Add fields → Set functions → Format → Sort
- Batch mode for configuring multiple fields

# excel_pivottable Tool

**Related tools**:
- excel_range - For source data ranges (use create-from-range)
- excel_table - For structured source tables (use create-from-table)
- excel_datamodel - For Data Model tables with DAX (use create-from-datamodel)

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
- create-from-datamodel enforces a 5 minute timeout; if the Data Model is stuck (privacy dialogs, refresh), you'll get `SuggestedNextActions` to resolve it instead of a hung session.

**Action disambiguation**:
- create-from-range: Create PivotTable from range address (sheetName + range parameters)
- create-from-table: Create PivotTable from Excel Table/ListObject (excelTableName = worksheet table name)
- create-from-datamodel: Create PivotTable from Power Pivot Data Model table (dataModelTableName = Data Model table name)
- add-row-field: Field goes to row area
- add-column-field: Field goes to column area
- add-value-field: Field goes to values area (calculations)
  - **OLAP Mode 1**: Add pre-existing DAX measure (fieldName = measure name or "[Measures].[Name]")
  - **OLAP Mode 2**: Auto-create DAX measure from column (fieldName = column name, specify aggregation function)
  - **Regular PivotTables**: Add column to values area with aggregation function
- add-filter-field: Field goes to filter area
- set-field-function: Change aggregation (Sum, Count, etc.)

**add-value-field for OLAP PivotTables**:
- **Pre-existing measure**: Use fieldName = "Total Sales" or "[Measures].[Total Sales]"
  - Adds existing measure without creating duplicate
  - aggregationFunction parameter ignored (measure formula defines aggregation)
  - Best for complex DAX measures with relationships, time intelligence, etc.
- **Auto-create from column**: Use fieldName = "Sales" (column name)
  - Creates new DAX measure: SUM('Table'[Sales]), COUNT('Table'[Column]), etc.
  - aggregationFunction parameter determines DAX function (Sum, Count, Average, etc.)
  - customName parameter sets measure name
  - Best for simple aggregations

**Parameter guidance for create actions**:
- create-from-range: sheetName, range, destinationSheet, destinationCell, pivotTableName
- create-from-table: excelTableName (Excel Table name), destinationSheet, destinationCell, pivotTableName
- create-from-datamodel: dataModelTableName (Data Model table name), destinationSheet, destinationCell, pivotTableName

**Common mistakes**:
- Creating PivotTable without source data → Prepare data first
- Not refreshing after source data changes → Use refresh action
- Wrong field type → Rows vs Columns vs Values have different purposes
- **OLAP**: Using column name when measure exists → Use measure name directly for existing measures
- **OLAP**: Creating duplicate measures → Check if measure exists first with excel_datamodel list-measures

**Workflow optimization**:
- Pattern: Create → Add fields → Set functions → Format → Sort
- **OLAP with existing measures**: list-fields → add-value-field (use measure name) → format
- **OLAP auto-create**: add-value-field (use column name + aggregation) → creates measure automatically

# excel_datamodel Tool

**⚠️ BEFORE CALLING 'create-measure'**: Check dax_measure.md elicitation for complete parameter checklist

**Related tools**:
- excel_powerquery - Load data to data model first (loadDestination='data-model' or 'both')
- excel_table - Use add-to-datamodel action to add Excel tables to data model
- excel_pivottable - Create PivotTables from data model tables
- excel_batch - Use for creating multiple measures (75-90% faster)

**Actions**: list-tables, view-table, list-columns, list-measures, view-measure, export-measure, create-measure, update-measure, delete-measure, list-relationships, create-relationship, update-relationship, delete-relationship, get-model-info, refresh

**When to use excel_datamodel**:
- DAX measures and calculated columns
- Table relationships
- Power Pivot Data Model operations
- Use excel_powerquery to load data to data model first
- Use excel_table for worksheet tables (not data model)

**Server-specific behavior**:
- Requires Power Query data loaded with loadDestination='data-model' or 'both'
- DAX formulas use Power Pivot syntax, not Excel worksheet formulas
- Measures created in data model, not in worksheets
- Relationships: One-to-many, many-to-one, one-to-one supported

**Action disambiguation**:
- list-tables: Show all tables in data model (not worksheet tables)
- view-table: Get detailed table info (columns, measures, row count)
- list-columns: List columns in a specific data model table
- list-measures: Show DAX measures only
- create-measure: Add new DAX calculation
- create-relationship: Link tables by columns
- refresh: Refresh all data model tables from sources

**Common mistakes**:
- Creating measures before loading data to model → Use loadDestination='data-model' first
- Confusing worksheet tables with data model tables → Different tools
- DAX syntax errors → Validate DAX before creating measures
- Not setting formatString on measures → Numbers display as general format

**Workflow optimization**:
- Multiple measures? Use begin_excel_batch (75-95% faster)
- Load queries to data model first: excel_powerquery(loadDestination='data-model')
- Then create measures: excel_datamodel(create-measure)
- Use display folders to organize measures

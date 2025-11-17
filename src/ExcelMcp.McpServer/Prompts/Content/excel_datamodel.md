# excel_datamodel - Server Quirks

**Action disambiguation**:
- list-tables: Data model tables only (NOT worksheet tables)
- read-table: Detailed info (columns, measures, row count, refresh date)
- list-columns: Columns in specific data model table
- create-measure vs DAX formulas: Measures are data model calculations, not worksheet formulas

**Server-specific quirks**:
- Requires data loaded with loadDestination='data-model' or 'both' first
- DAX syntax (Power Pivot) not Excel worksheet syntax
- Measures created in data model, NOT visible in worksheets until used in PivotTable
- formatString: Must set explicitly or numbers display as general format

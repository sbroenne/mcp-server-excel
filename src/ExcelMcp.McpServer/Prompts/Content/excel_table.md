# excel_table - Server Quirks

**Data Model workflow (CRITICAL)**:

Excel Tables on worksheets are NOT automatically in the Data Model (Power Pivot).
To analyze worksheet data with DAX measures:

1. Ensure data is formatted as an Excel Table (use create action if needed)
2. Use `add-to-datamodel` action to add the table to Power Pivot
3. Then use `excel_datamodel` to create DAX measures on it

**Action disambiguation**:

- create: Create NEW table from a range (requires sheetName, tableName, rangeAddress)
- read: Get table metadata (range, columns, style, row counts)
- get-data: Get actual table DATA as 2D array (use visibleOnly=true for filtered data)
- add-to-datamodel: Add an existing worksheet table to Power Pivot for DAX analysis
- append: Add rows to existing table (CSV format via style parameter)
- resize: Change table range (expand/contract)
- delete: Remove table (keeps data, removes table formatting)

**add-to-datamodel behavior**:

- Only works on Excel Tables (ListObjects), not plain ranges
- Table appears in Power Pivot with same name
- After adding, use excel_datamodel to create DAX measures
- Idempotent: calling on already-added table is a no-op

**When to use which tool**:

| Goal | Tool |
|------|------|
| Create/manage worksheet tables | excel_table |
| Add worksheet table to Power Pivot | excel_table (add-to-datamodel) |
| Import external data to Data Model | excel_powerquery (loadDestination='data-model') |
| Create DAX measures | excel_datamodel |
| Create PivotTables from Data Model | excel_pivottable |

**Common mistakes**:

- Trying to create DAX measures without first adding table to Data Model
- Using excel_datamodel to add tables (it only manages existing Data Model tables)
- Confusing get-data (returns cell values) with read (returns metadata)
- Forgetting hasHeaders parameter when creating tables from headerless data

**Server-specific quirks**:

- Style parameter is overloaded: table style name OR total function OR CSV data (context-dependent)
- visibleOnly parameter only applies to get-data action
- Table names must be unique within workbook (Excel requirement)

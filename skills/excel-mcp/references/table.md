# table - Server Quirks

**Data Model workflow (CRITICAL)**:

Excel Tables on worksheets are NOT automatically in the Data Model (Power Pivot).
To analyze worksheet data with DAX measures:

1. Ensure data is formatted as an Excel Table (use create action if needed)
2. Use `add-to-data-model` action to add the table to Power Pivot
3. Then use `datamodel` to create DAX measures on it

**Action disambiguation**:

- create: Create NEW table from a range (requires sheetName, tableName, rangeAddress). Pass `tableStyle` here to style at creation time.
- read: Get table metadata (range, columns, style, row counts)
- get-data: Get actual table DATA as 2D array (use visibleOnly=true for filtered data)
- rename: Rename an existing table
- delete: Remove table (keeps data, removes table formatting)
- resize: Change table range (expand/contract)
- set-style: Change table visual style (TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11). Default is TableStyleMedium2.
- toggle-totals: Show or hide the totals row (showTotals: true/false)
- set-column-total: Set the aggregate function on a totals-row column (Sum, Count, Average, Min, Max, None)
- add-to-data-model: Add an existing worksheet table to Power Pivot for DAX analysis
- append: Add rows to existing table (requires rows or rowsFile parameter)
- **create-from-dax**: Create table populated by a DAX EVALUATE query from Data Model
- **update-dax**: Update an existing DAX-backed table's query
- **get-dax**: Get the DAX query behind a DAX-backed table

**Table styling — always use table styles, not range_format**:

Excel Tables manage their own header/row/totals formatting through table styles. Never use `range_format(action: 'format-range')` on table header rows — it conflicts with the table style and produces inconsistent formatting.

| Goal | Correct approach |
|------|-----------------|
| Style a table | `table(action: 'set-style', tableStyle: 'TableStyleMedium2')` |
| Style at creation | `table(action: 'create', tableStyle: 'TableStyleMedium2', ...)` |
| Custom branding on table | Use a Medium/Dark table style that matches your palette — avoid overriding individual cells |

Common table style choices:
- `TableStyleMedium2` — standard blue, most widely used
- `TableStyleMedium9` — orange accent
- `TableStyleLight1` — minimal borders, no header fill
- `TableStyleDark1` — dark header with white text

**DAX-backed tables** (NEW):

Create worksheet tables populated by DAX EVALUATE queries against the Data Model.
Perfect for creating summary/report tables with aggregated data.

```
Workflow:
1. Have data in Data Model (via table add-to-data-model or powerquery)
2. Use create-from-dax with a DAX EVALUATE query
3. Table is created on worksheet with query results
4. Use update-dax to change the query, get-dax to inspect it
```

Example DAX queries for create-from-dax:
- `EVALUATE SUMMARIZE('Sales', 'Sales'[Region], "Total", SUM('Sales'[Amount]))`
- `EVALUATE TOPN(10, 'Products', 'Products'[Revenue], DESC)`
- `EVALUATE FILTER('Customers', 'Customers'[Country] = "USA")`

**add-to-data-model behavior**:

- Only works on Excel Tables (ListObjects), not plain ranges
- Table appears in Power Pivot with same name
- After adding, use datamodel to create DAX measures
- Idempotent: calling on already-added table is a no-op

**When to use which tool**:

| Goal | Tool |
|------|------|
| Create/manage worksheet tables | table |
| Add worksheet table to Power Pivot | table (add-to-data-model) |
| Import external data to Data Model | powerquery (loadDestination='data-model') |
| Create DAX measures | datamodel |
| Create PivotTables from Data Model | pivottable |

**Common mistakes**:

- Trying to create DAX measures without first adding table to Data Model
- Using datamodel to add tables (it only manages existing Data Model tables)
- Confusing get-data (returns cell values) with read (returns metadata)
- Forgetting hasHeaders parameter when creating tables from headerless data

**Server-specific quirks**:

- Style parameter is overloaded: table style name OR total function (context-dependent)
- csvData parameter: dedicated parameter for append action (CSV format: comma-separated, newline-separated rows)
- visibleOnly parameter only applies to get-data action
- Table names must be unique within workbook (Excel requirement)

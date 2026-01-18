# excel_datamodel - Server Quirks

**PREREQUISITE: Tables must be added to Data Model first!**

The Data Model (Power Pivot) only contains tables that were explicitly added.
You CANNOT create DAX measures on tables that aren't in the Data Model.

**How to add tables to the Data Model**:

| Source | Method |
|--------|--------|
| Worksheet Excel Table | excel_table with add-to-datamodel action |
| External file (CSV, etc.) | excel_powerquery with loadDestination='data-model' |
| Database/web source | excel_powerquery with loadDestination='data-model' |

**Automatic DAX Formatting**:

DAX formulas are automatically formatted on WRITE operations only (create-measure, update-measure) using the official Dax.Formatter library (SQLBI). Read operations (list-measures, read) return raw DAX as stored in Excel. Formatting adds ~100-500ms network latency per write operation but ensures consistent, professional code formatting. If formatting fails (network issues, API errors), the original DAX is saved unchanged - operations never fail due to formatting.

**Action disambiguation**:

- list-tables: List all tables currently in the Data Model
- list-measures: List all DAX measures (returns raw DAX from Excel)
- create-measure: Create a new DAX measure (DAX auto-formatted before saving)
- update-measure: Modify existing measure's formula/format/description (DAX auto-formatted before saving)
- delete-measure: Remove a measure
- delete-table: Remove table AND ALL its measures (DESTRUCTIVE!)
- read-info: Get Data Model metadata (culture, compatibility level)
- refresh: Refresh all Data Model data from sources
- **evaluate**: Execute DAX EVALUATE queries and return tabular results (read-only, no side effects)

**evaluate action** (NEW):

Execute any DAX EVALUATE query against the Data Model and return results as JSON.
Useful for ad-hoc analysis, testing DAX expressions, or extracting aggregated data.

```dax
// Examples of valid EVALUATE queries:
EVALUATE 'SalesTable'                    // Return entire table
EVALUATE TOPN(10, 'Sales', 'Sales'[Amount], DESC)  // Top 10 by amount
EVALUATE SUMMARIZE('Sales', 'Sales'[Region], "Total", SUM('Sales'[Amount]))  // Aggregation
EVALUATE FILTER('Products', 'Products'[Category] = "Electronics")  // Filtered
EVALUATE ROW("TotalRevenue", SUM('Sales'[Amount]))  // Single row result
```

**DAX measure creation**:

- tableName: Which table the measure belongs to (for organization)
- measureName: Display name for the measure
- daxFormula: DAX expression (e.g., "SUM(Sales[Revenue])")
- formatString: Optional number format (#,##0.00, 0%, $#,##0, etc.)

**Common DAX patterns**:

```dax
// Sum
SUM(TableName[ColumnName])

// Average
AVERAGE(TableName[ColumnName])

// Count rows
COUNTROWS(TableName)

// Calculated ratio
DIVIDE(SUM(Sales[Revenue]), SUM(Sales[Units]), 0)
```

**Common mistakes**:

- Creating measures before adding source table to Data Model → Error
- Using worksheet table names instead of Data Model table names
- Forgetting that delete-table removes ALL measures on that table
- Not specifying tableName when creating measures (required for organization)

**Server-specific quirks**:

- 2-minute auto-timeout on Data Model operations
- Table names in Data Model may differ from worksheet (check list-tables)
- Refresh refreshes ALL tables, not individual ones
- Measure names must be unique across entire Data Model (not per-table)

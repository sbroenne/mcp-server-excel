# excel_datamodel - Server Quirks

**PREREQUISITE: Tables must be added to Data Model first!**

The Data Model (Power Pivot) only contains tables that were explicitly added.
You CANNOT create DAX measures on tables that aren't in the Data Model.

## Excel Power Pivot Limitations (vs SSAS/Power BI)

| Feature | Power BI/SSAS | Excel Power Pivot | Workaround |
|---------|---------------|-------------------|------------|
| Calculated Tables | DAX: `MyTable = FILTER(...)` | NOT SUPPORTED | Use Power Query to create the table |
| Calculated Columns | DAX: `Table[Col] = ...` | NO COM API access | Use Power Query or DAX measures |
| Measures | Full support | Full support | - |
| Relationships | Full support | Full support | - |

**Key Insight**: Excel's COM API cannot create or modify calculated columns. If you need computed columns:
1. **Preferred**: Add the column in Power Query (computed at refresh time)
2. **Alternative**: Use a DAX measure instead (computed at query time)

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
- **execute-dmv**: Execute DMV queries for metadata discovery (SELECT * FROM $SYSTEM.*)

**evaluate action**:

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

**execute-dmv action** (DMV = Dynamic Management Views):

Execute SQL-like DMV queries to discover Data Model metadata.
DMVs are schema rowsets that expose Analysis Services internal information.

SYNTAX: `SELECT * FROM $SYSTEM.<SchemaRowset>`

IMPORTANT LIMITATIONS (Excel's embedded Analysis Services):
- ONLY `SELECT *` works - specific column selection (SELECT col1, col2) fails
- Some TMSCHEMA views return empty results despite Data Model having data
- Excel's embedded AS has limited support compared to full SQL Server Analysis Services

**Working DMV queries (verified in Excel):**

| DMV Query | Returns |
|-----------|---------|
| `SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES` | All DAX measures with formulas |
| `SELECT * FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS` | All relationships between tables |
| `SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY` | Calculation dependencies (useful for impact analysis) |
| `SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS` | Database/catalog metadata |
| `SELECT * FROM $SYSTEM.DISCOVER_SCHEMA_ROWSETS` | List all available DMVs |

**DMV queries that execute but may return empty in Excel:**

| DMV Query | Notes |
|-----------|-------|
| `SELECT * FROM $SYSTEM.TMSCHEMA_TABLES` | May return 0 rows in Excel's embedded AS |
| `SELECT * FROM $SYSTEM.TMSCHEMA_COLUMNS` | May return 0 rows in Excel's embedded AS |
| `SELECT * FROM $SYSTEM.TMSCHEMA_PARTITIONS` | May return 0 rows in Excel's embedded AS |

**Full list of TMSCHEMA DMVs** (from MS-SSAS-T protocol):

| Category | DMVs |
|----------|------|
| Model Structure | TMSCHEMA_MODEL, TMSCHEMA_TABLES, TMSCHEMA_COLUMNS, TMSCHEMA_HIERARCHIES, TMSCHEMA_LEVELS |
| Measures/KPIs | TMSCHEMA_MEASURES, TMSCHEMA_KPIS, TMSCHEMA_FORMAT_STRING_DEFINITIONS |
| Relationships | TMSCHEMA_RELATIONSHIPS |
| Security | TMSCHEMA_ROLES, TMSCHEMA_ROLE_MEMBERSHIPS, TMSCHEMA_TABLE_PERMISSIONS, TMSCHEMA_COLUMN_PERMISSIONS |
| Partitions | TMSCHEMA_PARTITIONS, TMSCHEMA_DATA_SOURCES |
| Metadata | TMSCHEMA_ANNOTATIONS, TMSCHEMA_EXTENDED_PROPERTIES, TMSCHEMA_CULTURES, TMSCHEMA_OBJECT_TRANSLATIONS |
| Perspectives | TMSCHEMA_PERSPECTIVES, TMSCHEMA_PERSPECTIVE_TABLES, TMSCHEMA_PERSPECTIVE_COLUMNS, TMSCHEMA_PERSPECTIVE_MEASURES |
| Calculations | TMSCHEMA_CALCULATION_GROUPS, TMSCHEMA_CALCULATION_ITEMS, TMSCHEMA_EXPRESSIONS |

**DISCOVER DMVs** (server/analysis metadata):

| DMV | Description |
|-----|-------------|
| DISCOVER_CALC_DEPENDENCY | Dependencies between objects (great for impact analysis) |
| DISCOVER_SCHEMA_ROWSETS | List all available schema rowsets |
| DISCOVER_PROPERTIES | Server properties |
| DISCOVER_KEYWORDS | Reserved keywords |
| DISCOVER_LITERALS | Supported literals |

**Example use cases:**

```sql
-- Find all measures and their DAX formulas
SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES

-- Discover what objects a measure depends on
SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY

-- List all relationships
SELECT * FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS

-- Get catalog information
SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS
```

Reference: [Microsoft DMV Documentation](https://learn.microsoft.com/en-us/analysis-services/instances/use-dynamic-management-views-dmvs-to-monitor-analysis-services)

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

## Star Schema Architecture

**Why use DAX over Power Query for calculations?**

- DAX recalculates on refresh without re-running Power Query
- Useful when lookup/rate tables change frequently

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

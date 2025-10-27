using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts for Power Query development patterns and best practices.
/// </summary>
[McpServerPromptType]
public static class ExcelPowerQueryPrompts
{
    /// <summary>
    /// Comprehensive Power Query M language reference and common patterns.
    /// </summary>
    [McpServerPrompt(Name = "excel_powerquery_mcode_reference")]
    [Description("M language reference with common Power Query patterns and functions")]
    public static ChatMessage MCodeReference()
    {
        return new ChatMessage(ChatRole.User, @"# Power Query M Language Reference

## Common M Functions

### Data Source Functions
- **Excel.Workbook()** - Load Excel files
- **Csv.Document()** - Load CSV files
- **Web.Contents()** - Fetch web data
- **Json.Document()** - Parse JSON
- **Sql.Database()** - Connect to SQL Server

### Transformation Functions
- **Table.SelectRows()** - Filter rows with predicate
- **Table.SelectColumns()** - Keep only specified columns
- **Table.RemoveColumns()** - Remove columns
- **Table.RenameColumns()** - Rename columns
- **Table.TransformColumnTypes()** - Change data types
- **Table.AddColumn()** - Add calculated column
- **Table.ExpandRecordColumn()** - Unnest nested records
- **Table.Group()** - Group and aggregate

### Common Patterns

#### Pattern 1: Load CSV with Type Conversion
```m
let
    Source = Csv.Document(File.Contents(""C:\data\sales.csv""), [Delimiter="","", Encoding=65001]),
    Headers = Table.PromoteHeaders(Source),
    Types = Table.TransformColumnTypes(Headers, {
        {""Date"", type date},
        {""Amount"", type number},
        {""Category"", type text}
    })
in
    Types
```

#### Pattern 2: Filter and Transform
```m
let
    Source = Excel.CurrentWorkbook(){[Name=""RawData""]}[Content],
    Filtered = Table.SelectRows(Source, each [Status] = ""Active""),
    Added = Table.AddColumn(Filtered, ""Total"", each [Quantity] * [Price]),
    Sorted = Table.Sort(Added, {{""Total"", Order.Descending}})
in
    Sorted
```

#### Pattern 3: Merge Queries (Join)
```m
let
    Orders = Excel.CurrentWorkbook(){[Name=""Orders""]}[Content],
    Customers = Excel.CurrentWorkbook(){[Name=""Customers""]}[Content],
    Merged = Table.NestedJoin(Orders, {""CustomerID""}, Customers, {""ID""}, ""CustomerInfo"", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, ""CustomerInfo"", {""Name"", ""Email""})
in
    Expanded
```

#### Pattern 4: Parameterized Query
```m
let
    // Reference named range parameter
    StartDate = Excel.CurrentWorkbook(){[Name=""StartDate""]}[Content]{0}[Column1],
    Source = Sql.Database(""server"", ""database""),
    Filtered = Table.SelectRows(Source, each [OrderDate] >= StartDate)
in
    Filtered
```

## Privacy Levels

When combining data from multiple sources, specify privacy level:
- **None** - Ignore privacy (fastest, least secure)
- **Private** - Prevent data sharing (most secure, recommended for sensitive data)
- **Organizational** - Share within organization
- **Public** - Public data sources only

## Best Practices

1. **Always specify types** - Use Table.TransformColumnTypes() early
2. **Filter early** - Apply filters before transformations
3. **Use Table.Buffer()** - Cache expensive operations
4. **Avoid hardcoded paths** - Use parameters for file paths
5. **Name steps clearly** - Use descriptive variable names
6. **Test incrementally** - Build query step by step

## Common Errors

### Error: ""Expression.Error: The key didn't match any rows""
**Fix:** Check table/column names, ensure they exist

### Error: ""DataSource.Error: Web.Contents failed""
**Fix:** Check URL, network connectivity, authentication

### Error: ""Formula.Firewall: Query references other queries""
**Fix:** Set appropriate privacy level with --privacy-level parameter");
    }

    /// <summary>
    /// Guide for managing Power Query connection settings and refresh behavior.
    /// </summary>
    [McpServerPrompt(Name = "excel_powerquery_connections")]
    [Description("Power Query connection management and refresh configuration")]
    public static ChatMessage ConnectionManagement()
    {
        return new ChatMessage(ChatRole.User, @"# Power Query Connection Management

## Load Configurations

Power Query supports 3 load configurations:

### 1. Load to Table (Default)
Creates a worksheet table with query results.
```bash
excel_powerquery set-load-to-table --query SalesData --sheet Results
```

### 2. Load to Data Model
Loads data into Power Pivot (in-memory database).
```bash
excel_powerquery set-load-to-data-model --query SalesData
```

### 3. Connection Only
Query exists but doesn't load data (used by other queries).
```bash
excel_powerquery set-connection-only --query HelperQuery
```

## Refresh Behavior

### Manual Refresh
```bash
excel_powerquery refresh --query SalesData
```

### Refresh All Queries
When queries depend on each other, refresh in dependency order:
1. Connection-only queries first
2. Dependent queries after
3. Power Pivot last (if using data model)

## Privacy Levels

Required when combining data sources:

### When to Use Each Level

**Private** - Use for:
- Corporate databases
- Files with sensitive data
- Internal APIs
- Default for unknown data

**Organizational** - Use for:
- Shared corporate data sources
- Internal data warehouses
- Department-level data

**Public** - Use for:
- Public APIs
- Open data sources
- Published datasets

**None** - Use for:
- Development/testing only
- Single data source queries
- Performance-critical scenarios (no privacy checks)

## Example Workflow

```bash
# 1. Import connection-only helper query (no privacy needed)
excel_powerquery import --query DateDimension --file dates.pq --connection-only

# 2. Import main query that uses helper (requires privacy level)
excel_powerquery import --query Sales --file sales.pq --privacy-level Private

# 3. Set load configuration
excel_powerquery set-load-to-table --query Sales --sheet SalesData

# 4. Refresh
excel_powerquery refresh --query Sales
```

## Troubleshooting

### Error: ""Privacy levels are required""
**Solution:** Add --privacy-level parameter to import/update

### Error: ""Unable to refresh""
**Solution:** Check data source connectivity, credentials

### Query is slow
**Solutions:**
1. Filter data at source (SQL WHERE clause, not M filter)
2. Remove unnecessary columns early
3. Use Table.Buffer() for expensive operations
4. Consider connection-only for intermediate steps");
    }

    /// <summary>
    /// Common Power Query development workflows and patterns.
    /// </summary>
    [McpServerPrompt(Name = "excel_powerquery_workflows")]
    [Description("Step-by-step workflows for common Power Query development scenarios")]
    public static ChatMessage DevelopmentWorkflows()
    {
        return new ChatMessage(ChatRole.User, @"# Power Query Development Workflows

## Workflow 1: Version-Controlled Query Development

**Goal:** Develop queries in .pq files, track in Git

### Steps
1. **Export existing query (if migrating)**
```bash
excel_powerquery export --file report.xlsx --query SalesData --output queries/sales.pq
```

2. **Edit M code** in your favorite editor (VS Code recommended)

3. **Import updated query**
```bash
excel_powerquery import --file report.xlsx --query SalesData --source queries/sales.pq
```

4. **Commit to Git**
```bash
git add queries/sales.pq
git commit -m ""Add sales data query with date filter""
```

### Batch Session Pattern (Faster for Multiple Changes)
```typescript
const { batchId } = await begin_excel_batch({ filePath: ""report.xlsx"" });

await excel_powerquery({ batchId, action: ""import"", queryName: ""Sales"", mCodeFile: ""sales.pq"" });
await excel_powerquery({ batchId, action: ""import"", queryName: ""Products"", mCodeFile: ""products.pq"" });
await excel_powerquery({ batchId, action: ""import"", queryName: ""Customers"", mCodeFile: ""customers.pq"" });

await commit_excel_batch({ batchId, save: true });
```

## Workflow 2: Data Source Migration (Dev → Prod)

**Goal:** Switch query from dev database to production

### Steps
1. **Export query from dev workbook**
```bash
excel_powerquery export --file dev-report.xlsx --query SalesData --output sales.pq
```

2. **Edit connection string** in sales.pq
```m
// Before
Source = Sql.Database(""dev-server"", ""dev-database"")

// After  
Source = Sql.Database(""prod-server"", ""prod-database"")
```

3. **Import to production workbook**
```bash
excel_powerquery import --file prod-report.xlsx --query SalesData --source sales.pq --privacy-level Organizational
```

## Workflow 3: Incremental Data Loading

**Goal:** Load only new/changed data since last refresh

### Steps
1. **Create parameter for last refresh date**
```bash
excel_parameter create --file report.xlsx --name LastRefresh --reference Sheet1!A1
excel_parameter set --file report.xlsx --name LastRefresh --value ""2024-01-01""
```

2. **Create query that uses parameter**
```m
let
    LastRefresh = Excel.CurrentWorkbook(){[Name=""LastRefresh""]}[Content]{0}[Column1],
    Source = Sql.Database(""server"", ""database""),
    FilteredRows = Table.SelectRows(Source, each [ModifiedDate] >= LastRefresh),
    Now = DateTime.LocalNow()
in
    FilteredRows
```

3. **After refresh, update parameter**
```bash
excel_powerquery refresh --file report.xlsx --query IncrementalData
excel_parameter set --file report.xlsx --name LastRefresh --value ""2024-01-15""
```

## Workflow 4: Multi-Source Data Combination

**Goal:** Combine data from Excel, CSV, and SQL

### Steps
1. **Create individual queries (connection-only)**
```bash
# Excel source
excel_powerquery import --query ExcelData --file excel-source.pq --connection-only

# CSV source
excel_powerquery import --query CsvData --file csv-source.pq --connection-only

# SQL source
excel_powerquery import --query SqlData --file sql-source.pq --connection-only
```

2. **Create combine query** that merges all sources
```m
let
    Excel = ExcelData,
    Csv = CsvData,
    Sql = SqlData,
    Combined = Table.Combine({Excel, Csv, Sql}),
    Deduped = Table.Distinct(Combined, {""ID""})
in
    Deduped
```

3. **Import with privacy level**
```bash
excel_powerquery import --query Combined --file combine.pq --privacy-level Private
```

## Workflow 5: Query Refactoring

**Goal:** Optimize slow query

### Steps
1. **Export current query**
```bash
excel_powerquery export --file report.xlsx --query SlowQuery --output slow.pq
```

2. **Benchmark current performance** (note refresh time)

3. **Refactor M code** (apply best practices)
   - Move filters to data source
   - Remove unnecessary columns early
   - Use Table.Buffer() for repeated access
   - Simplify transformations

4. **Update query**
```bash
excel_powerquery update --file report.xlsx --query SlowQuery --source optimized.pq
```

5. **Test refresh** and compare performance

6. **Commit optimized version**

## Best Practices

1. **One query per .pq file** - Easier to track changes
2. **Descriptive filenames** - Match query names
3. **Test before commit** - Always refresh after import
4. **Document dependencies** - Note which queries use parameters
5. **Use batch sessions** - For multi-query updates
6. **Version control** - Track all .pq files in Git");
    }
}

# Data & Analytics Features

Import, transform, model, and summarize data with Power Query, DAX, Excel Tables, PivotTables, and external data connections.

[← Back to the complete feature reference](../../FEATURES.md)

---

## 🔄 Power Query & M Code (12 operations)

Import, transform, and refresh data with Power Query. Every operation is a single-call atomic workflow.

**Discovery:**
- **List:** Return compact query metadata, exact load state, and an M preview of at most 80 characters; full M is never included
- **View:** View one query's full M code and exact load state
- **Get Load Config:** Get current load configuration

**Lifecycle:**
- **Create:** Import + load in one operation (atomic workflow), preserving M code by default
- **Update:** Update M code, preserving M code by default, with optional auto-refresh
- **Rename:** Rename a Power Query (trim + case-insensitive uniqueness check)
- **Unload:** Remove data from all destinations (keeps query definition)
- **Delete:** Remove Power Query from workbook

**Loading & Refresh:**
- **Refresh:** Refresh a Power Query with timeout detection
- **Refresh All:** Batch refresh all queries in workbook
- **Load To:** Configure load destination and refresh (atomic)

**Advanced:**
- **Evaluate:** Execute M code directly and return results (without creating a permanent query)

**Notes:**
- **M-code formatting:** M code is preserved exactly by default. Create and Update can opt in to remote formatting with `formatMCode=true`, which sends M code to powerqueryformatter.com and adds network latency. If remote formatting fails, the original M code is saved unchanged.
- **Inline or file input:** Required M, DAX, and DMV text accepts exactly one inline value or its matching file alias (`mCodeFile`, `daxFormulaFile`, `daxQueryFile`, or `dmvQueryFile` in batch JSON; snake_case in MCP; kebab-case options in the CLI). Optional `update-measure` DAX input may omit both forms; whenever a value is supplied, inline and file forms remain mutually exclusive. Files must exist and be readable.
- **Timeout representation:** Public CLI, batch, and MCP timeouts are integer seconds. Power Query refresh/refresh-all accepts 0–2147483; omission or `0` uses the 30-minute data-operation default. Connection, Data Model, PivotTable, and VBA timeouts accept 1–2147483.
- **Load destinations:** `worksheet`, `data-model`, `both`, and `connection-only` are case-insensitive aliases for the corresponding generated enum values. Unknown values are rejected before any load destination is changed.
- **Action-specific parameters:** CLI and MCP schemas expose the category-wide parameter union, but each action rejects explicitly supplied parameters it does not use. In particular, `load-to` rejects `timeout` and uses the fixed 30-minute data-operation timeout.
- **Truthful reads:** List, View, and Get Load Config share exact worksheet/Data Model-aware load detection. List fails explicitly if Excel cannot inspect a query rather than silently omitting it.
- **Exact query identity:** Load detection, refresh, load-to, unload, and delete compare the parsed mashup `Location` exactly and case-insensitively, so prefix names such as `A` and `AA` cannot affect each other.
- **Evaluate cleanup:** Temporary queries, worksheets, tables, and generically named workbook connections are removed by exact mashup identity. Cleanup failures return an actionable error and never report success.

---

## 📊 Data Model & DAX (Power Pivot) (20 operations)

Build a Power Pivot Data Model — manage tables, DAX measures, and relationships, then query it.

**Tables & Columns:**
- **List Tables:** Discover all tables in the Data Model
- **Read Table:** Get specific table information
- **Rename Table:** Rename a Data Model table (best-effort via Power Query; returns clear error if not supported)
- **Delete Table:** Remove table from Data Model
- **List Columns:** List columns for a table

**Measures:**
- **List Measures:** List all DAX measures with formula previews
- **Read Measure:** Get one measure's full DAX formula, format, description, and table
- **Create Measure:** Create new DAX measure, preserving DAX by default
- **Update Measure:** Modify existing measure, preserving DAX by default
- **Delete Measure:** Remove measure from model

**Relationships:**
- **List Relationships:** View all table relationships
- **Read Relationship:** Get specific relationship info
- **Create Relationship:** Create relationship between tables
- **Update Relationship:** Modify relationship (toggle active/inactive)
- **Delete Relationship:** Remove relationship

**Model & Queries:**
- **Read Info:** Get comprehensive model information
- **Read Connection:** Inspect the embedded model connection, command type, and connected table names
- **Refresh:** Refresh entire Data Model
- **Evaluate:** Execute DAX EVALUATE queries and return tabular results (for ad-hoc analysis)
- **Execute DMV:** Execute SQL-like DMV (Dynamic Management View) queries for metadata discovery

**Notes:**
- **DAX formatting:** DAX formulas are preserved exactly by default, subject to Excel locale separator translation. CreateMeasure and UpdateMeasure can opt in to remote formatting with `formatDax=true`, which sends DAX to daxformatter.com and adds network latency. If remote formatting fails, the original DAX is saved unchanged.
- **Measure formats:** Create Measure and Update Measure accept General, Currency, Decimal, Percentage, and WholeNumber case-insensitively. Unknown values are rejected before Excel is called. Create defaults to General when the format is omitted or empty; Update keeps the existing format when omitted or empty.
- **Workbook connections:** Use the existing `connection list` action to list workbook connections and Power Query sources; this is not a Data Model action.
- **Source metadata:** Read Table returns each table's source connection name, description, type, and model-membership flag.
- **COM limitations:** Excel exposes calculated columns as read-only entries but provides no reliable PIA formula/mutation or live refresh-status API. Use Power Query for computed columns.

---

## 📇 Excel Tables (ListObjects) (27 operations)

Create and manage Excel Tables (ListObjects) — structured ranges with styling, filtering, and sorting.

**Lifecycle:**
- **List:** List Excel Tables in a worksheet or workbook
- **Read:** Get table structure (columns, range, style)
- **Create:** Create a new Excel Table from a range
- **Rename:** Rename an existing table
- **Resize:** Resize table range to match new data bounds
- **Delete:** Remove a table (keeps underlying cell data)

**Styling & Formatting:**
- **Apply Style:** Apply a built-in table style
- **Toggle Totals Row:** Show/hide the totals row
- **Set Column Totals:** Configure per-column total function (Sum, Average, Count, etc.)

**Data Operations:**
- **Append Rows:** Add rows to the end of a table
- **Get Table Data:** Read table data as a 2D array, with optional visible-only filtering
- **Add to Data Model:** Load a table into the Power Pivot Data Model

**DAX-Backed Tables:**
- **Create from DAX:** Create an Excel Table populated by a DAX EVALUATE query
- **Update DAX:** Change the DAX query of an existing DAX-backed table
- **Get DAX:** Retrieve the DAX query info from a table

**Filter Operations:**
- **Apply Filter (Criteria):** Filter a column using comparison criteria
- **Apply Filter (Values):** Filter a column to a specific set of values
- **Clear Filters:** Remove all active filters
- **Get Filter State:** Read current filter criteria

**Column Management:**
- **Add Column:** Insert a new column
- **Remove Column:** Delete a column
- **Rename Column:** Rename a column header

**Structured References:**
- **Get Structured Reference:** Get formula syntax for a table column or range

**Sorting:**
- **Sort (Single Column):** Sort by one column
- **Sort (Multi-Column):** Sort by up to 3 columns/levels

**Number Formatting:**
- **Get Column Number Formats:** Read number formats applied to columns
- **Set Column Number Formats:** Apply number formats to columns

---

## 📈 PivotTables (35 operations)

Create and configure PivotTables from ranges, Excel Tables, or the Data Model.

**Creation:**
- **Create from Range:** Build a PivotTable from a cell range
- **Create from Excel Table:** Build a PivotTable from an Excel Table
- **Create from Data Model:** Build an OLAP PivotTable from the Data Model

**Field Management:**
- **List Fields:** List all fields across row, column, value, and filter areas
- **Add Row Field / Column Field / Value Field / Filter Field:** Add a field to the given area
- **Remove Field:** Remove a field from the PivotTable

**Field Configuration:**
- **Set Field Function:** Set aggregation function (Sum, Average, Count, Min, Max, etc.)
- **Set Field Name:** Set a custom display name for a field
- **Set Field Number Format:** Apply a number format to a value field
- **Set Field Filter:** Apply filter criteria to a field
- **Sort Field:** Sort a field ascending/descending

**Grouping:**
- **Group by Date / Number:** Build automatic date hierarchies or numeric bands
- **Group Items:** Combine selected visible items into a named manual group
- **Ungroup Field:** Remove manual grouping and restore the original field

**Calculated Fields (Regular PivotTables):**
- **List Calculated Fields:** List calculated fields on a regular PivotTable
- **Create Calculated Field:** Add a calculated field
- **Delete Calculated Field:** Remove a calculated field

**Calculated Members (OLAP/Data Model PivotTables):**
- **List Calculated Members:** List calculated members on an OLAP PivotTable
- **Create Calculated Member:** Add a calculated member
- **Delete Calculated Member:** Remove a calculated member

**Layout & Formatting:**
- **Set Layout:** Switch between table and outline layout
- **Set Subtotals Display:** Show/hide subtotals
- **Set Grand Totals Display:** Show/hide grand totals

**Data Operations:**
- **Get PivotTable Data:** Read PivotTable data as a 2D array
- **Refresh:** Refresh the PivotTable from its source
- **Drill Through:** Expand a regular PivotTable value cell into its underlying source rows

**PivotCache Configuration:**
- **Get Cache Options:** Read refresh, retained-item, optimization, and saved-source settings
- **Set Cache Options:** Configure supported regular-cache options; unsupported OLAP/OLE DB mutations are rejected

**Lifecycle:**
- **List:** List PivotTables in a worksheet or workbook
- **Read:** Get PivotTable info
- **Delete:** Remove a PivotTable

---

## 🔌 Data Connections (11 operations)

Create and refresh external OLEDB/ODBC data connections.

**Operations:**
- **List:** View all data connections
- **View:** Get connection details
- **Create:** Create OLEDB/ODBC connections (requires provider installed)
- **Test:** Verify connection validity
- **Refresh:** Refresh connection data
- **Get Refresh Status:** Report whether a typed OLEDB/ODBC refresh is active
- **Cancel Refresh:** Cancel an active typed OLEDB/ODBC refresh
- **Delete:** Remove connection
- **Load To:** Load connection data to worksheet (when supported)
- **Get Properties:** Get connection string and metadata
- **Set Properties:** Update connection string, command text, and settings

**Notes:**
- **Supported types:** OLEDB (requires Microsoft.ACE.OLEDB.16.0 or similar), ODBC (requires ODBC driver installed), and Power Query connections (atomic redirect to `powerquery`).
- **Text/web imports:** Use `querytable` for direct local text/CSV or legacy HTML imports; use `powerquery` for transformations and modern connectors.
- **Safe cleanup:** Connection delete and load-to remove only QueryTables owned by the exact WorkbookConnection; similarly named QueryTables are preserved.

---

## 🌐 QueryTables (9 operations)

Manage local worksheet QueryTables through the typed Excel PIA.

**Operations:**
- **List / View:** Discover QueryTables and read destination, source type, and refresh configuration
- **Create Text:** Import text or CSV files with delimiter, qualifier, encoding, and header settings
- **Create Web:** Import legacy HTML pages or selected tables through Excel's web-query engine
- **Set Properties:** Configure background refresh, refresh-on-open, refresh period, sizing, and formatting preservation
- **Refresh / Get Refresh Status / Cancel Refresh:** Control and inspect refresh execution
- **Delete:** Remove a QueryTable

**Boundaries:** QueryTables do not expose Power Query M, modern cloud connectors, sharing, presence, mentions, reactions, or coauthoring state.

---

## Related feature areas

- [Cells & workbooks](CELLS-WORKBOOKS.md) — prepare ranges, formulas, worksheets, and files for analysis
- [Charts & visualization](CHARTS-VISUALS.md) — present analytical results with charts, slicers, and formatting
- [Automation & advanced](AUTOMATION-ADVANCED.md) — extend workflows with VBA, Python, and What-If Analysis
- [Example workflows](../USE-CASES.md) — see these capabilities combined in practical requests
- [Installation](../INSTALLATION.md) — choose and configure the MCP Server or CLI

## Task guides

- [Refresh Power Query from an AI assistant](../guides/REFRESH-POWER-QUERY.md)
- [Query the Excel Data Model with DAX](../guides/QUERY-DATA-MODEL-WITH-DAX.md)
- [Build and update PivotTables with an AI assistant](../guides/AUTOMATE-PIVOTTABLES.md)

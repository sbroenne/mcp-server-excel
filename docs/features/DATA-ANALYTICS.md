# Data & Analytics Features

Bring data into Excel, transform it, connect related tables, and build summaries
that use Excel's own query and calculation engines.

[Back to the feature overview](../../FEATURES.md)

These are capability summaries. Current command details come from CLI help or
MCP tool descriptions; the linked guides explain practical decisions and workflows.

---

## Power Query & M Code (12 operations)

- **Import and transform:** Create or update reusable M queries and load their results to worksheets, the Data Model, or both.
- **Manage queries:** Inspect definitions and load destinations, rename queries, or remove their loads while retaining the query.
- **Refresh data:** Refresh selected queries or attempt a workbook-wide query refresh, with errors reported rather than hidden.
- **Try M code:** Evaluate code without retaining a permanent query.

Creating or loading a query can execute its data sources. Evaluation also runs
code and temporarily changes the workbook, so it is not a read-only audit.
Failed loads can leave objects behind; inspect what survived before retrying.
Definition-only staging queries cannot refresh independently, and a batch
refresh failure does not roll back loads that already finished.

M code is preserved unless remote formatting is requested. Remote formatting
sends code to an external service and needs consent.

[Loading and recovery guidance](../reference/powerquery.md) |
[Refresh walkthrough](../guides/REFRESH-POWER-QUERY.md)

---

## Data Model & DAX (Power Pivot) (20 operations)

- **Inspect the model:** Discover tables, columns, measures, relationships, and the embedded model connection.
- **Build calculations:** Create and update DAX measures with meaningful formats and descriptions.
- **Connect tables:** Manage relationships between detail data and unique lookup tables.
- **Query results:** Evaluate DAX for analytical results or inspect model metadata with supported DMV queries.
- **Maintain the model:** Refresh data and manage tables and measures.

A worksheet Table is not automatically part of the Data Model. After changing
a worksheet source, refresh the model before relying on its results; refresh
dependent PivotTables afterward. Power Query refresh updates its model load.
DAX queries need the Microsoft Analysis Services OLE DB provider.

Calculated-column editing and DAX calculated-table creation are not exposed
through Excel's supported automation here. Use Power Query for stored
transformations and DAX measures for analytical calculations. Remote DAX
formatting sends code outside the computer and needs consent.

Excel can store a measure definition without evaluating it. Readback proves
storage, not a valid calculation; evaluate a query referencing the measure to
check its result. Format readback reports Excel's actual type and available
properties, distinguishing Decimal from Percentage.

[Model and calculation guidance](../reference/datamodel.md) |
[DAX walkthrough](../guides/QUERY-DATA-MODEL-WITH-DAX.md)

---

## Excel Tables (ListObjects) (27 operations)

- **Structure data:** Create, inspect, rename, resize, and append to worksheet Tables; check conversion risks before creating one.
- **Manage columns:** Add, remove, or rename columns and use structured references in formulas.
- **Present and summarize:** Apply Table styles, number formats, and column totals.
- **Filter and sort:** Use Excel's native value, date, comparison, color, and other supported filters, with single- or multi-column sorting.
- **Connect to analysis:** Add worksheet data to the model or display a DAX query as a worksheet Table.

Reuse a suitable existing Table rather than rebuilding it. Creation checks can
identify unsafe headers, merged cells, and sorting risks, but warnings are not
a guarantee that conversion suits the workbook's layout. Removing the Table
keeps cell data but can affect dependent objects. A DAX-backed worksheet Table
is a displayed query result, not a calculated table inside the model.

[Table workflow guidance](../reference/table.md)

---

## PivotTables (45 operations)

- **Build summaries:** Create PivotTables from ranges, worksheet Tables, or the Data Model, then configure rows, columns, values, and filters.
- **Choose calculations:** Set aggregation and Show Values As independently; use calculated fields for regular PivotTables or supported calculated members for model-backed ones.
- **Explore data:** Filter, sort, group, expand or collapse supported items, and drill into regular PivotTable source rows.
- **Control presentation:** Use Compact, Tabular, or Outline layouts, repeated labels, styles, subtotals, and grand totals.
- **Maintain sources:** Inspect refresh/cache settings and shared users, or change a supported source without rebuilding unrelated PivotTables.

Regular and Data Model PivotTables have different capabilities. Native grouping,
calculated filters, item expansion, and drill-through described here are
regular-PivotTable features. Use model columns and measures where those native
features are unavailable.

Creating a PivotTable does not configure its fields. Read the actual results
after field changes and source refreshes. Shared caches and connected slicers
also constrain source changes.

[PivotTable workflow guidance](../reference/pivottable.md) |
[PivotTable walkthrough](../guides/AUTOMATE-PIVOTTABLES.md)

---

## Data Connections (11 operations)

- **Connect existing sources:** Create, inspect, and maintain supported OLEDB or ODBC workbook connections.
- **Refresh and troubleshoot:** Test access, refresh data, inspect active refresh state, or cancel a supported refresh.
- **Manage loads:** Load supported connections to worksheets or remove their associated load objects.

The appropriate provider or driver must be installed. Power Query connections
use Power Query behavior rather than ordinary OLEDB/ODBC connection handling.
Connection details can contain credentials: do not publish them in reports or
error summaries. Use direct text/web imports for simple imports and Power Query
for transformations and modern connectors.

[Choosing an import workflow](../reference/querytable.md) |
[Safe access and error handling](../reference/behavioral-rules.md)

---

## QueryTables (9 operations)

- **Import files:** Bring local text or CSV data into a worksheet with control over how Excel interprets it.
- **Import legacy web data:** Load HTML pages or selected tables through Excel's web-query engine.
- **Maintain imports:** Inspect destinations and refresh settings, refresh or cancel loads, and remove QueryTables.

Legacy web queries are not a general browser or modern cloud connector.
QueryTables do not expose Power Query M or Microsoft 365 sharing, presence,
or collaboration state.

[Text and web import guidance](../reference/querytable.md)

---

## Related feature areas

- [Cells & workbooks](CELLS-WORKBOOKS.md) - prepare cells, formulas, sheets, and files
- [Charts & visualization](CHARTS-VISUALS.md) - present results with charts and interactive filters
- [Automation & advanced](AUTOMATION-ADVANCED.md) - extend workflows with code and What-If Analysis
- [Example workflows](../USE-CASES.md)

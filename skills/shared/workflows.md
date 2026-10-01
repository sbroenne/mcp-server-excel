# Excel MCP Server - Key Constraints

These sequencing workflows describe the **Windows COM backend**. Power Query,
Data Model/DAX, charts, and screenshots are unavailable in the experimental
macOS beta. Basic worksheet/range/named-range setup can be used on Mac, but does
not enable a Power Query workflow. Check
[macOS support](https://excelmcpserver.dev/macos-support/) before selecting actions.

## Excel Power Pivot Limitations

Excel's Power Pivot has key limitations compared to Power BI/SSAS:

| Feature | Availability | Workaround |
|---------|--------------|------------|
| Calculated Tables | NOT SUPPORTED | Create table in Power Query |
| Calculated Columns | No COM API | Use Power Query or DAX measures |
| Measures | Full support | - |
| Relationships | Full support | - |

**Implication**: Design your architecture to put computed columns in Power Query, not DAX.

## Architecture: Power Query vs DAX

| Layer | Use For | Update Frequency |
|-------|---------|------------------|
| Power Query | Data loading, transformations, computed columns | When source changes |
| Relationships | Star schema structure | Rarely |
| DAX | Business calculations, aggregations | Frequently |

Power Query prepares stored model data; DAX evaluates measures in the query's
filter context. Refresh source data when it changes before relying on DAX results.

## Tool Sequencing

### Data Model Prerequisites

Load the required tables into the Data Model, or add existing worksheet Tables
to it, before creating measures: `powerquery load-to` with
`load_destination: 'data-model'` (MCP) / `--load-destination data-model` (CLI),
or `table add-to-data-model`. Add relationships only when a calculation needs
cross-table filtering; a measure over one table does not need a relationship.
Discover existing model tables and relationships before creating new ones.
Use `datamodel_relationship` (MCP) / `datamodelrelationship` (CLI) with
`create-relationship`, then `datamodel create-measure` when needed.
See [Data Model guidance](datamodel.md) for native examples.

### Power Query Development Lifecycle
For authorized query development, use `powerquery evaluate`, then
`powerquery create` or `powerquery update` for the intended query. Create loads its chosen
destination; update refreshes by default.
Use `powerquery load-to` when changing destinations and `powerquery refresh`
when loaded data needs updating.
Prefer evaluation for new or changed code; trivial or already-validated code with
unchanged dependencies does not need redundant evaluation. Persisting untested
code can leave a broken query in the workbook. Evaluation uses temporary workbook
objects and executes M code; it is not a read-only audit operation.
See [Power Query](powerquery.md).

### Parameter Setup for Power Query
When parameter setup is requested, reuse the intended cells and named reference.
Use `worksheet create` (MCP) / `sheet create` (CLI) only if a setup sheet is
needed, `range set-values` for the parameter values, and `namedrange create`
for the named reference to those cells.
Power Query reads via `Excel.CurrentWorkbook(){[Name = "..."]}`

## Verification Commands

Use `powerquery list`, `powerquery view`, and `powerquery get-load-config`
for stored definitions and load configuration; those do not
prove loaded values are current. Read the affected loaded data when that is the
requested result. Use `datamodel list-measures` and `datamodel evaluate` for
measures, `datamodel_relationship` (MCP) / `datamodelrelationship` (CLI) with
`list-relationships` for relationship metadata, and `chart read` for chart
series/bounds. Screenshots are
useful for layout when an interactive desktop is available, not for every task.

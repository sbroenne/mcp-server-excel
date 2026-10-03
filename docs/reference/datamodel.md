# Data Model and DAX

Worksheet Tables and Data Model tables are separate. Add an existing Excel Table
to the model, or load a Power Query to `data-model`/`both`, before creating
relationships or measures. Use model table listings to discover exact names.

## Creating a measure

The following assumes an existing worksheet Table named `Sales` with an `Amount`
column and a captured session. Check each result before continuing.

```mcp
table(action: 'add-to-data-model', session_id: sessionId, table_name: 'Sales')
datamodel(action: 'create-measure', session_id: sessionId, table_name: 'Sales', measure_name: 'Total Sales', dax_formula: 'SUM(Sales[Amount])')
datamodel(action: 'evaluate', session_id: sessionId, dax_query: 'EVALUATE ROW("Total", [Total Sales])')
```

```cli
excelcli -q table add-to-data-model --session $sessionId --table-name Sales
excelcli -q datamodel create-measure --session $sessionId --table-name Sales --measure-name 'Total Sales' --dax-formula 'SUM(Sales[Amount])'
excelcli -q datamodel evaluate --session $sessionId --dax-query 'EVALUATE ROW("Total", [Total Sales])'
```

Measure names are unique across the model, not just their home table. Use update
for an existing measure. Formats are General, Currency, Decimal, Percentage, or
WholeNumber. On create, an omitted format defaults to General; on update it keeps
the existing format. Failure to obtain an explicitly requested format fails the
operation; it does not silently substitute General. Update obtains that format
before changing the formula or description. A later Excel write failure can
still leave partial changes; no rollback is promised.
Measure readback identifies the actual Excel format interface, so Decimal and
Percentage remain distinct even though both expose decimal-place and separator
properties. Failed format-property reads do not silently report General.
Supply native DAX with comma argument separators and decimal
points; create and update pass it to Excel without regional separator rewriting.
Remote formatting requires explicit consent and remains off by default:
MCP `format_dax: true` or CLI `--format-dax true`.

On the tested decimal-comma Excel installation, native measure writes still reject
some numeric arguments followed by commas, including `DATEADD(..., -1, MONTH)`
and `IF(..., 1.5, 0)`, even when the same DAX evaluates successfully as a query.
Report the Excel error instead of rewriting the supplied formula, changing
regional settings, or enabling remote formatting as a workaround.

## Refresh is not calculation

After worksheet Table edits or appends, refresh the model before querying it.
Power Query refresh updates its model load. Refresh dependent PivotTables after
their source is current.

DAX measures are evaluated at query time in the current filter context; they do
not reload source data. A changed external rate/lookup table still needs a source
refresh. Put stored transformations and computed columns in Power Query; use
measures for calculations over current model data.

Refresh can target the whole model or one table and accepts a caller timeout.
Other model query operations have their own limits (including a two-minute DMV
query timeout); do not apply one blanket timeout to every action. Excel exposes
no reliable live model-refresh status or per-table refresh timestamp here.
Refresh failures preserve Excel's error details rather than assuming that
model-level refresh is unsupported. Check the reported source or engine error
before choosing recovery; refreshing individual tables is not a universal fix.

Model metadata reads fail when a required property cannot be read. A reported
row count of zero means Excel returned zero, not that its property getter failed.

## Relationships and unsupported features

Relationships connect the detail/many side to the unique lookup/one side with
compatible column types. Only one relationship per table pair can be active;
use DAX `USERELATIONSHIP` for an inactive one. List relationships before changing
or deleting model tables. Deleting a model table also removes its measures and
relationships.

Excel Power Pivot does not expose DAX calculated-table creation through these
tools. Its COM column API exposes names and types, not calculated-column formula
reading or mutation. Use Power Query computed columns or measures instead.

## Querying and displaying results

`evaluate` runs DAX `EVALUATE` queries. `execute-dmv` reads model metadata using
SQL-like schema-rowset queries. Both require the Microsoft Analysis Services
OLE DB provider (MSOLAP). If it is missing, report the prerequisite; see
[Microsoft's client libraries](https://learn.microsoft.com/analysis-services/client-libraries).
Do not diagnose every Excel error as missing MSOLAP.

Use [DMV guidance](dmv-reference.md) for supported rowsets and Excel limitations.
Some rowsets return no rows even when a model exists.

| Desired result | Approach |
|----------------|----------|
| Query results in a worksheet | DAX-backed Table via `create-from-dax` |
| Query results returned to the caller | DAX `evaluate` |
| Interactive grouping and filtering | Model-backed PivotTable |
| Chart linked to PivotTable fields | Verified live PivotChart |

For per-row revenue, use `SUMX(Sales, Sales[Quantity]*Sales[UnitPrice])`, not a
product of column totals. Use `DIVIDE` when a denominator may be zero.

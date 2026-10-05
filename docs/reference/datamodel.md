# Data Model and DAX

Worksheet Tables and Data Model tables are separate. Add an existing Excel Table
to the model, or load a Power Query to `data-model`/`both`, before creating
relationships or measures. Use model table listings to discover exact names.
Current commands and inputs come from CLI help or MCP tool descriptions.

## Creating a measure

The following assumes an existing worksheet Table named `Sales` with an `Amount`
column and a captured session. Check each result before continuing.

```mcp
table(action: 'add-to-data-model', workbook_session_id: sessionId, table_name: 'Sales')
datamodel(action: 'create-measure', workbook_session_id: sessionId, table_name: 'Sales', measure_name: 'Total Sales', dax_formula: 'SUM(Sales[Amount])')
datamodel_read(action: 'evaluate', workbook_session_id: sessionId, dax_query: 'EVALUATE ROW("Total", [Total Sales])')
```

```cli
excelcli -q table add-to-data-model --session $sessionId --table-name Sales
excelcli -q datamodel create-measure --session $sessionId --table-name Sales --measure-name 'Total Sales' --dax-formula 'SUM(Sales[Amount])'
excelcli -q datamodel evaluate --session $sessionId --dax-query 'EVALUATE ROW("Total", [Total Sales])'
```

Measure names are unique across the model, not just their home table. Update
an existing measure rather than trying to create it again. Choose its display
format from the result's meaning; percentages and decimals are not interchangeable.
Read back the actual definition and result, especially after failure, because
a write can leave partial changes.

Supply native DAX with comma argument separators and decimal points. Remote
formatting is optional, sends code to an external service, and requires consent.

When Windows uses a comma as the decimal mark, Excel misreads a comma that
touches a number in a measure formula, for example turning `IF(..., 1.5, 0)`
into an invalid formula or `DATEADD(..., -1, MONTH)` into `-1.`. Create and
update then add one space on each side of such commas (`IF(..., 1.5 , 0)`,
`DATEADD(..., -1 , MONTH)`). Text, quoted names, column references, and comments
are left alone. The DAX meaning is unchanged, and the result `message` says
spaces were added. Do not change regional settings or enable remote formatting
as a workaround. Measure readback returns the DAX that
the model stores, with decimal points, including after the workbook is reopened.

No machine or Excel regional settings are changed.
Measure writes store a definition; Excel may defer validation of an unknown
function until evaluation. A successful write is not proof of a successful
calculation. Evaluate the measure in the intended context and inspect its value.

## Refresh is not calculation

After worksheet Table edits or appends, refresh the model before querying it.
Power Query refresh updates its model load. Refresh dependent PivotTables after
their source is current.

DAX measures are evaluated at query time in the current filter context; they do
not reload source data. A changed external rate/lookup table still needs a source
refresh. Put stored transformations and computed columns in Power Query; use
measures for calculations over current model data.

Refresh and querying have different completion limits; use their current help
rather than applying one blanket timeout to every operation. Excel exposes
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

DAX queries return analytical results; DMV queries inspect model metadata.
Both require the Microsoft Analysis Services
OLE DB provider (MSOLAP). If it is missing, report the prerequisite; see
[Microsoft's client libraries](https://learn.microsoft.com/analysis-services/client-libraries).
Do not diagnose every Excel error as missing MSOLAP.
Date columns returned by `evaluate` may appear as Excel date-serial numbers,
not formatted date strings. Check the returned values and types before using
them as display-ready dates.

Use [DMV guidance](dmv-reference.md) for supported rowsets and Excel limitations.
Some rowsets return no rows even when a model exists.

| Desired result | Approach |
|----------------|----------|
| Query results in a worksheet | DAX-backed worksheet Table |
| Query results returned to the caller | DAX query evaluation |
| Interactive grouping and filtering | Model-backed PivotTable |
| Chart linked to PivotTable fields | Verified live PivotChart |

For per-row revenue, use `SUMX(Sales, Sales[Quantity]*Sales[UnitPrice])`, not a
product of column totals. Use `DIVIDE` when a denominator may be zero.

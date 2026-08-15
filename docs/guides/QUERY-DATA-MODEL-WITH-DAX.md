# Query the Excel Data Model with DAX

Excel's Data Model (Power Pivot) is a real tabular analytics engine living inside
the workbook. ExcelMcp gives an AI assistant direct access to it: create measures,
run DAX queries, inspect metadata with DMVs, and manage relationships.

This is only possible because ExcelMcp drives the actual Excel application. The
Data Model is not something you can read or write by parsing the file.

## What you ask for

> Add a `Total Revenue` measure to the Sales table in `analysis.xlsx` and show me
> revenue by region.

> List every measure in this workbook and its DAX formula.

## Prerequisite for querying: MSOLAP

Creating measures and relationships works out of the box. **Executing DAX queries
and DMVs additionally requires the MSOLAP OLE DB provider**, which is not part of
Excel.

The simplest way to get it is to install **Power BI Desktop** (free), which ships
the provider. Alternatively install the Microsoft Analysis Services OLE DB
provider directly.

Without MSOLAP you can still create, update, and list measures — you just cannot
run `evaluate` or `execute-dmv`.

## Load data into the model

Two routes:

```powershell
# From a worksheet Excel Table
excelcli -q table add-to-data-model --session $session --table-name SalesTable

# Or from Power Query - the preferred route for external data
excelcli -q powerquery create --session $session --query-name Sales `
  --m-code-file .\sales.m --load-destination data-model
```

Power Query is preferred: it gives you refreshable, typed, transformable data, and
refreshing the query auto-syncs the model.

## Create measures

```powershell
excelcli -q datamodel create-measure --session $session `
  --table-name SalesTable --measure-name "Total Revenue" `
  --dax-formula "SUMX(SalesTable, SalesTable[Quantity] * SalesTable[UnitPrice])"
```

Measures are reusable across every PivotTable in the workbook, work across related
tables, and support time intelligence — none of which PivotTable calculated fields
can do.

## Run DAX queries

```powershell
excelcli -q datamodel evaluate --session $session `
  --dax-query "EVALUATE SUMMARIZECOLUMNS(SalesTable[Region], \"Revenue\", [Total Revenue])"
```

This is the fastest way to sanity-check a measure before wiring it into a
PivotTable.

## Inspect the model with DMVs

Dynamic Management Views expose the model's own metadata:

```powershell
excelcli -q datamodel execute-dmv --session $session --dmv-query "SELECT * FROM `$SYSTEM.TMSCHEMA_MEASURES"
excelcli -q datamodel execute-dmv --session $session --dmv-query "SELECT * FROM `$SYSTEM.TMSCHEMA_RELATIONSHIPS"
excelcli -q datamodel execute-dmv --session $session --dmv-query "SELECT * FROM `$SYSTEM.TMSCHEMA_COLUMNS"
```

DMVs are the reliable way to discover what is actually in the model, including
objects the COM object model does not expose.

## Relationships

```powershell
excelcli -q datamodel create-relationship --session $session `
  --from-table Sales --from-column ProductId `
  --to-table Products --to-column Id
```

Relationships require matching data types on both sides. This is the most common
cause of "relationship could not be created" — a date stored as text on one side,
or an ID typed as number on one side and text on the other. Setting explicit types
in Power Query (`Table.TransformColumnTypes`) prevents it.

## Keep the model in sync

The worksheet table and the Data Model table are **separate copies** of the data.

```text
table(action: 'append', ...)      # worksheet only
datamodel(action: 'refresh')      # now the model matches
pivottable(action: 'refresh')     # now the pivot matches
```

Refreshing a Power Query that loads to the model does all of this for you.

## Known limitations

**No calculated tables or calculated columns.** Excel's Power Pivot COM API does
not expose them, so ExcelMcp cannot create them. Do that work in Power Query
instead — add the column there and reload. This is an Excel limitation, not an
ExcelMcp one; Power BI Desktop supports both.

**DMV queries only support `SELECT *`.** Column projection, `WHERE`, and `ORDER BY`
are not supported by the Excel DMV endpoint. Filter the results after retrieval.

**Hidden Data Model objects are invisible to COM.** Objects marked hidden in the
model do not appear through the object model — use DMVs to see them.

**Measure names are case-sensitive in DAX references** and must be unique across
the whole model, not just per table.

## Related

- [Power Query, DAX and analytics operations](../features/DATA-ANALYTICS.md)
- [Build and update PivotTables with an AI assistant](AUTOMATE-PIVOTTABLES.md)
- [Refresh Power Query from an AI assistant](REFRESH-POWER-QUERY.md)

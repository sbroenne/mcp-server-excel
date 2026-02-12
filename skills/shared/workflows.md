# Excel MCP Server - Key Constraints

These are the critical constraints and workarounds specific to Excel automation via COM.

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

**Why separate?** DAX measures recalculate on refresh without re-running Power Query. Useful when lookup/rate tables change often.

## Tool Sequencing

### Data Model Prerequisites
```
1. Load table (powerquery refresh loadDestination="data-model")
2. THEN create relationships (datamodel_relationship with create-relationship action)
3. THEN create measures (datamodel create-measure)
```
Skipping step 1 causes "table not found" errors.

### Power Query Development Lifecycle
```
1. powerquery evaluate (test M code without persisting - catches errors early)
2. powerquery create/update (store validated query in workbook)
3. powerquery refresh/load-to (load data to destination)
```
Skipping step 1 causes broken queries in workbook and cryptic COM errors.

### Parameter Setup for Power Query
```
1. worksheet create (e.g., "_Setup")
2. range set-values (parameter values)
3. namedrange create (named reference)
```
Power Query reads via `Excel.CurrentWorkbook(){[Name = "..."]}`

## Verification Commands

```
After Power Query: powerquery list, powerquery view
After refresh:     datamodel list-tables
After measure:     datamodel list-measures, datamodel evaluate
After relationship: datamodel_relationship list-relationships
```

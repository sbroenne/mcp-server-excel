# pivottable - Server Quirks

## CRITICAL: Required Parameters

**`pivot_table_name` is REQUIRED for almost all PivotTable operations** across `pivottable`, `pivottable_calc`, and `pivottable_field` tools. The only exception is `list` (which lists all PivotTables). Always specify the PivotTable name.

## Calculated Fields vs DAX Measures

PivotTable calculated fields work well for simple single-table formulas. Use DAX measures for complex scenarios.

| Feature | PivotTable Calculated Field | DAX Measure |
|---------|----------------------------|-------------|
| Single-table formulas | ✅ Works (e.g., `=Qty*Price`) | ✅ Works |
| Cross-table | NOT SUPPORTED | Full support |
| Complex logic | Limited | Full DAX |
| Reusable | Per PivotTable only | Across all PivotTables |

### Calculated Field Workflow

```
pivottable_calc(create-calculated-field, pivot_table_name="SalesPivot", field_name="Revenue", formula="=Quantity*UnitPrice")
pivottable_field(add-value-field, pivot_table_name="SalesPivot", field_name="Revenue", aggregation_function="Sum")
```

### DAX Measure Workflow (for complex scenarios)

```
table(add-to-data-model, table_name="Sales")
datamodel(create-measure, table_name="Sales", measure_name="Revenue", dax_formula="SUMX(Sales, Sales[Quantity]*Sales[UnitPrice])")
pivottable(create-from-datamodel, pivot_table_name="SalesPivot", destination_sheet="Analysis", destination_cell="A3", table_name="Sales")
```

### When to Use DAX Instead of Calculated Fields

- Multi-table calculations (need relationships between tables)
- Complex logic (time intelligence, YTD, running totals)
- Calculations involving filtered contexts
- Reusable measures across multiple PivotTables

## PivotTable Source Types

| Source | Create Action | Supports DAX Measures? |
|--------|---------------|------------------------|
| Worksheet Table | `create-from-table` | NO - worksheet PivotTable |
| Data Model | `create-from-datamodel` | YES - full DAX support |
| External | `create-from-range` with `source_range` | NO |

**Rule**: If you need calculated revenue/aggregations, use Data Model as source.

## Refresh Behavior (CRITICAL)

PivotTables do NOT auto-refresh when source data changes!

**After adding rows to source table:**
```
table(append, ...)           # Add rows to worksheet table
pivottable(refresh, ...)     # Refresh PivotTable to see new rows
datamodel(refresh)           # ALSO refresh Data Model if using DAX measures
```

**After Power Query refresh:**
```
powerquery(refresh, ...)     # Refreshes Power Query AND Data Model
# PivotTables connected to Data Model auto-refresh
```

## PivotCache Options

- `pivottable(get-cache-options)`: Read refresh, retained-item, optimization, and saved-source settings.
- `pivottable(set-cache-options)`: Set `enable_refresh`, `refresh_on_file_open`, `missing_items_limit`, `optimize_cache`, or `save_source_data`.
- `missing_items_limit` values: `Default`, `None`, `Max`, `Max2`.
- Deleted-item retention applies only to regular PivotTables. OLAP/Data Model caches manage members in the model.

## Field Configuration

### Row/Column/Value Fields

When creating PivotTables, configure fields in order:
1. Add Row fields: `pivottable_field(add-row-field, pivot_table_name="SalesPivot", field_name="Region")`
2. Add Column fields: `pivottable_field(add-column-field, pivot_table_name="SalesPivot", field_name="Year")`
3. Add Value fields: `pivottable_field(add-value-field, pivot_table_name="SalesPivot", field_name="Amount", aggregation_function="Sum")`
4. Add filters: `pivottable_field(add-filter-field, pivot_table_name="SalesPivot", field_name="Status")`
5. **Refresh to update display**: `pivottable(refresh, pivot_table_name="SalesPivot")`

**IMPORTANT**: Field operations are structural only - they modify the PivotTable layout but don't trigger visual refresh. Call `pivottable(refresh)` after configuring all fields to update the display. This is especially important for OLAP/Data Model PivotTables.

### Manual Grouping

```
pivottable_field(group-items, pivot_table_name="SalesPivot", field_name="Region", item_names=["North", "South"], group_name="Core Regions")
# Use groupedFieldName from the result:
pivottable_field(ungroup-field, pivot_table_name="SalesPivot", grouped_field_name="Region2")
```

Manual grouping requires a regular PivotTable and a field already placed in the Row or Column area. OLAP/Data Model PivotTables must add grouping columns in the model.

### Drill Through

```
pivottable(drill-through, pivot_table_name="SalesPivot", cell_address="G4")
```

The target must be a value cell in a regular PivotTable data body. Excel creates a new worksheet containing the underlying source rows. OLAP/Data Model drill-through is provider-dependent and intentionally not exposed as a deterministic operation.

### Aggregation Functions for Value Fields

| Function | Use Case |
|----------|----------|
| Sum | Totals (revenue, quantity) |
| Count | Record counts |
| Average | Mean values |
| Min/Max | Extremes |
| CountNums | Count numbers only |
| StdDev/Var | Statistical analysis |

## Common Patterns

### Revenue Analysis from Worksheet Table

```
# Option 1: Add revenue column to source table FIRST
range(set-formulas, sheet_name="Sales", range_address="I2", formulas=[["=[@Quantity]*[@UnitPrice]"]])
pivottable(create-from-table, pivot_table_name="SalesPivot", destination_sheet="Analysis", destination_cell="A3", table_name="SalesTable")
pivottable_field(add-value-field, pivot_table_name="SalesPivot", field_name="Revenue", aggregation_function="Sum")

# Option 2: Use Data Model (RECOMMENDED)
table(add-to-data-model, table_name="SalesTable")
datamodel(create-measure, table_name="SalesTable", measure_name="Revenue", dax_formula="SUMX(SalesTable, SalesTable[Quantity]*SalesTable[UnitPrice])")
pivottable(create-from-datamodel, pivot_table_name="SalesPivot", destination_sheet="Analysis", destination_cell="A3", table_name="SalesTable")
```

### Multi-Table Analysis

Always use Data Model for multi-table analysis:
```
table(add-to-data-model, table_name="Sales")
table(add-to-data-model, table_name="Products")
datamodel_relationship(create-relationship, from_table="Sales", from_column="ProductID", to_table="Products", to_column="ProductID")
datamodel(create-measure, table_name="Sales", measure_name="Revenue", dax_formula="SUMX(Sales, RELATED(Products[Price])*Sales[Quantity])")
pivottable(create-from-datamodel, pivot_table_name="SalesPivot", destination_sheet="Analysis", destination_cell="A3", table_name="Sales")
```

## Layout Styles

The `row_layout` parameter on `pivottable_calc(set-layout)` controls PivotTable appearance:

| Value | Style | Description |
|-------|-------|-------------|
| 0 | Compact | Default, nested row labels |
| 1 | Tabular | Each field in separate column, best for exports |
| 2 | Outline | Hierarchical with expand/collapse |

## Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| "Unknown field" aggregation error | Calculated field type limitation | Use DAX measure instead |
| "Table not found" | Source not in Data Model | Add with `table(add-to-data-model)` |
| "Field not found" | Typo or Data Model not refreshed | Refresh Data Model, check field names |
| Data doesn't update | Source changed without refresh | Call `pivottable(refresh)` |
| DAX measures missing | Created on worksheet PivotTable | Use `create-from-datamodel` |

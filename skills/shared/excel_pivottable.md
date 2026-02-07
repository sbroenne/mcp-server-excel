````markdown
# excel_pivottable - Server Quirks

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
excel_pivottable_calc(CreateCalculatedField, fieldName="Revenue", formula="=Quantity*UnitPrice")
excel_pivottable_field(AddValueField, fieldName="Revenue", aggregationFunction="Sum")
```

### DAX Measure Workflow (for complex scenarios)

```
excel_table(add-to-datamodel, tableName="Sales")
excel_datamodel(create-measure, measureName="Revenue", daxFormula="SUMX(Sales, Sales[Quantity]*Sales[UnitPrice])")
excel_pivottable(create-from-datamodel, ...)  # Measure automatically available
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
| External | `create` with sourceRange | NO |

**Rule**: If you need calculated revenue/aggregations, use Data Model as source.

## Refresh Behavior (CRITICAL)

PivotTables do NOT auto-refresh when source data changes!

**After adding rows to source table:**
```
excel_table(append, ...)           # Add rows to worksheet table
excel_pivottable(refresh, ...)     # Refresh PivotTable to see new rows
excel_datamodel(refresh)           # ALSO refresh Data Model if using DAX measures
```

**After Power Query refresh:**
```
excel_powerquery(refresh, ...)     # Refreshes Power Query AND Data Model
# PivotTables connected to Data Model auto-refresh
```

## Field Configuration

### Row/Column/Value Fields

When creating PivotTables, configure fields in order:
1. Add Row fields: `excel_pivottable_field(AddRowField, fieldName="Region")`
2. Add Column fields: `excel_pivottable_field(AddColumnField, fieldName="Year")`  
3. Add Value fields: `excel_pivottable_field(AddValueField, fieldName="Amount", aggregationFunction="Sum")`
4. Add filters: `excel_pivottable_field(AddFilterField, fieldName="Status")`
5. **Refresh to update display**: `excel_pivottable(refresh, pivotTableName="...")`

**IMPORTANT**: Field operations are structural only - they modify the PivotTable layout but don't trigger visual refresh. Call `excel_pivottable(refresh)` after configuring all fields to update the display. This is especially important for OLAP/Data Model PivotTables.

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
excel_range(set-formula, sheetName="Sales", rangeAddress="I2", formula="=[@Quantity]*[@UnitPrice]")
excel_pivottable(create-from-table, sourceTableName="SalesTable", ...)
excel_pivottable_field(AddValueField, fieldName="Revenue", aggregationFunction="Sum")  # Works!

# Option 2: Use Data Model (RECOMMENDED)
excel_table(add-to-data-model, tableName="SalesTable")
excel_datamodel(create-measure, measureName="Revenue", daxFormula="SUMX(SalesTable, SalesTable[Quantity]*SalesTable[UnitPrice])")
excel_pivottable(create-from-datamodel, ...)  # Measure automatically available
```

### Multi-Table Analysis

Always use Data Model for multi-table analysis:
```
excel_table(add-to-data-model, tableName="Sales")
excel_table(add-to-data-model, tableName="Products")
excel_datamodel_rel(create-relationship, fromTable="Sales", fromColumn="ProductID", toTable="Products", toColumn="ProductID")
excel_datamodel(create-measure, tableName="Sales", measureName="Revenue", daxFormula="SUMX(Sales, RELATED(Products[Price])*Sales[Quantity])")
excel_pivottable(create-from-datamodel)
```

## Layout Styles

The `layoutStyle` parameter controls PivotTable appearance:

| Value | Style | Description |
|-------|-------|-------------|
| 0 | Compact | Default, nested row labels |
| 1 | Outline | Each field in separate column |
| 2 | Tabular | Flat table format, best for exports |

## Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| "Unknown field" aggregation error | Calculated field type limitation | Use DAX measure instead |
| "Table not found" | Source not in Data Model | Add with `excel_table(add-to-data-model)` |
| "Field not found" | Typo or Data Model not refreshed | Refresh Data Model, check field names |
| Data doesn't update | Source changed without refresh | Call `excel_pivottable(refresh)` |
| DAX measures missing | Created on worksheet PivotTable | Use `create-from-datamodel` |

````

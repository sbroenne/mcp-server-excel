````markdown
# pivottable - Server Quirks

## CRITICAL: Required Parameters

**`pivotTableName` is REQUIRED for almost all PivotTable operations** across `pivottable`, `pivottable_calc`, and `pivottable_field` tools. The only exception is `list` (which lists all PivotTables). Always specify the PivotTable name.

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
pivottable_calc(CreateCalculatedField, fieldName="Revenue", formula="=Quantity*UnitPrice")
pivottable_field(AddValueField, fieldName="Revenue", aggregationFunction="Sum")
```

### DAX Measure Workflow (for complex scenarios)

```
table(add-to-data-model, tableName="Sales")
datamodel(create-measure, measureName="Revenue", daxFormula="SUMX(Sales, Sales[Quantity]*Sales[UnitPrice])")
pivottable(create-from-datamodel, ...)  # Measure automatically available
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
table(append, ...)           # Add rows to worksheet table
pivottable(refresh, ...)     # Refresh PivotTable to see new rows
datamodel(refresh)           # ALSO refresh Data Model if using DAX measures
```

**After Power Query refresh:**
```
powerquery(refresh, ...)     # Refreshes Power Query AND Data Model
# PivotTables connected to Data Model auto-refresh
```

## Field Configuration

### Row/Column/Value Fields

When creating PivotTables, configure fields in order:
1. Add Row fields: `pivottable_field(AddRowField, fieldName="Region")`
2. Add Column fields: `pivottable_field(AddColumnField, fieldName="Year")`  
3. Add Value fields: `pivottable_field(AddValueField, fieldName="Amount", aggregationFunction="Sum")`
4. Add filters: `pivottable_field(AddFilterField, fieldName="Status")`
5. **Refresh to update display**: `pivottable(refresh, pivotTableName="...")`

**IMPORTANT**: Field operations are structural only - they modify the PivotTable layout but don't trigger visual refresh. Call `pivottable(refresh)` after configuring all fields to update the display. This is especially important for OLAP/Data Model PivotTables.

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
range(set-formula, sheetName="Sales", rangeAddress="I2", formula="=[@Quantity]*[@UnitPrice]")
pivottable(create-from-table, sourceTableName="SalesTable", ...)
pivottable_field(AddValueField, fieldName="Revenue", aggregationFunction="Sum")  # Works!

# Option 2: Use Data Model (RECOMMENDED)
table(add-to-data-model, tableName="SalesTable")
datamodel(create-measure, measureName="Revenue", daxFormula="SUMX(SalesTable, SalesTable[Quantity]*SalesTable[UnitPrice])")
pivottable(create-from-datamodel, ...)  # Measure automatically available
```

### Multi-Table Analysis

Always use Data Model for multi-table analysis:
```
table(add-to-data-model, tableName="Sales")
table(add-to-data-model, tableName="Products")
datamodel_relationship(create-relationship, fromTable="Sales", fromColumn="ProductID", toTable="Products", toColumn="ProductID")
datamodel(create-measure, tableName="Sales", measureName="Revenue", daxFormula="SUMX(Sales, RELATED(Products[Price])*Sales[Quantity])")
pivottable(create-from-datamodel)
```

## Layout Styles

The `layoutStyle` parameter controls PivotTable appearance:

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

````

---
name: excel-mcp
description: >
  Automate Microsoft Excel on Windows via COM interop. Use when creating, reading, 
  or modifying Excel workbooks. Supports Power Query (M code), Data Model (DAX measures), 
  PivotTables, Tables, Ranges, Charts, Slicers, Formatting, VBA macros, and connections.
  Triggers: Excel, spreadsheet, workbook, xlsx, Power Query, DAX, PivotTable, VBA.
license: MIT
version: 1.3.0
tags:
  - excel
  - automation
  - mcp
  - windows
  - powerquery
  - dax
  - pivottable
  - vba
  - data-model
  - charts
  - formatting
  - slicers
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel MCP Server Skill

Server-specific guidance for Excel MCP Server. Tools are auto-discovered - this documents quirks, workflows, and gotchas.

## Preconditions

- Windows host with Microsoft Excel installed (2016+)
- Use full Windows paths: `C:\Users\Name\Documents\Report.xlsx`
- Excel files must not be open in another Excel instance
- VBA operations require "Trust access to VBA project object model" enabled in Excel Trust Center

## CRITICAL: Execution Rules (MUST FOLLOW)

**NEVER ask clarifying questions.** Make reasonable assumptions and execute immediately.

**Common scenarios where you MUST NOT ask - just do it:**

| Instead of Asking | DO THIS |
|-------------------|---------|
| "Which file should I use?" | Call `excel_file(list)` to discover open sessions |
| "Where should I create the PivotTable?" | Create a new sheet with descriptive name |
| "What table should I use?" | Call `excel_table(list)` to discover tables |
| "Should I add this to the Data Model?" | Yes, if DAX measures are needed |
| "What format should I use?" | Use professional defaults (currency, dates, etc.) |

**ALWAYS call tools instead of explaining:**

| DON'T | DO |
|-------|-----|
| "I would use excel_slicer to filter..." | CALL `excel_slicer(set-slicer-selection)` |
| "You can use excel_datamodel to create measures..." | CALL `excel_datamodel(create-measure)` |
| "The revenue calculation would be..." | CALL the tool, then REPORT the actual result |

**When uncertain:**
1. List first: `excel_file(list)`, `excel_table(list)`, `excel_worksheet(list)`
2. Read metadata: `excel_table(read)`, `excel_datamodel(list-tables)`
3. Then execute the operation with discovered information

## Tool Selection Quick Reference

**Which tool for which task:**

| Task | Tool | Key Action |
|------|------|------------|
| Create/open/save workbooks | `excel_file` | open, create, close |
| Write/read cell data | `excel_range` | set-values, get-values |
| Create tables from data | `excel_table` | create |
| Add data to existing table | `excel_table` | append (use csvData param) |
| Add table to Power Pivot | `excel_table` | add-to-datamodel |
| Create DAX formulas | `excel_datamodel` | create-measure |
| Query Data Model | `excel_datamodel` | evaluate |
| Create PivotTables | `excel_pivottable` | create, create-from-datamodel |
| Filter with slicers | `excel_slicer` | set-slicer-selection |
| Create charts | `excel_chart` | create-from-range |
| Import external data | `excel_powerquery` | create, refresh |

**Multi-step workflows - FOLLOW THIS ORDER:**

```
Data Analysis:
excel_range(set-values) → excel_table(create) → excel_table(add-to-datamodel) → excel_datamodel(create-measure) → excel_pivottable(create-from-datamodel)

Filtering:
excel_slicer(create-slicer) → excel_slicer(set-slicer-selection)

External Data:
excel_powerquery(create) → excel_powerquery(refresh) → [Data Model ready]
```

## Session Workflow

1. **Open/Create**: 
   - **Existing file**: `excel_file(open, excelPath)` → returns `sessionId`
   - **New file**: `excel_file(create, excelPath)` → creates file AND returns `sessionId` in single operation
2. **Perform Operations**: Pass `sessionId` to all tool calls
3. **Save and Close**: `excel_file(close, sessionId, save=true)` to persist changes

**Session Tips:**
- Call `excel_file(list)` first to check for existing sessions (reuse if file already open)
- Use `create` for new files (single Excel startup, faster performance)
- Check `canClose=true` before closing (no active operations running)
- Without `save=true`, all changes are discarded

**Session Timeout:**
- Default: 5 minutes per operation. Use `timeoutSeconds` to customize (range: 10-3600)
- For long operations (large Power Query refresh, complex DAX): `excel_file(open, excelPath, timeoutSeconds=600)`
- Timeout triggers aggressive cleanup - Excel may be in inconsistent state after timeout

## Format Cells After Setting Values (CRITICAL)

Without formatting, dates appear as serial numbers and currency as plain numbers.

| Data Type | Format Code | Example Result |
|-----------|-------------|----------------|
| Currency (USD) | `$#,##0.00` | $1,234.56 |
| Currency (EUR) | `€#,##0.00` | €1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |
| Number | `#,##0.00` | 1,234.56 |

**Workflow**: `excel_range(set-values)` → `excel_range_format(set-number-format)`

## Format Data as Excel Tables (CRITICAL)

Always convert tabular data to Excel Tables, not plain ranges:

```
1. excel_range(set-values)  # Write data with headers
2. excel_table(create)      # Convert to Table
```

**Benefits**: Structured references, auto-expand, built-in filtering, required for Data Model.

## Core Rules

1. **2D Arrays**: Values and formulas use 2D arrays. Single cell = `[[value]]`
2. **Targeted Updates**: Modify specific cells, not entire structures
3. **List Before Delete**: Verify names exist before delete/rename operations
4. **US Format Codes**: Use US locale format codes (`#,##0.00` not `#.##0,00`)
5. **Check Results**: Always check `success` and `errorMessage` in responses
6. **Follow Suggestions**: Act on `suggestedNextActions` when provided

## CRITICAL: Error Recovery (NEVER ASK - TRY ALTERNATIVES)

**When a tool returns an error, TRY AN ALTERNATIVE APPROACH. Never ask the user what to do.**

| Error Message | DON'T Ask | DO This Instead |
|---------------|-----------|-----------------|
| "Data Model returns old values" | "Should I refresh?" | Call `excel_datamodel(refresh)` then retry |
| "Table not found in Data Model" | "Should I add it?" | Call `excel_table(add-to-datamodel)` first |
| "Field not found" | "What field name is correct?" | Call `excel_table(read)` to discover fields |
| "Query already exists" | "Should I update it?" | Use `excel_powerquery(update)` instead of `create` |

**Error recovery patterns:**

```
# Pattern 1: DAX query returns stale data after table changes
ERROR: DAX evaluate returns old values
FIX:
  excel_datamodel(refresh)
  excel_datamodel(evaluate, daxQuery="...")

# Pattern 2: Measure creation fails - table not in Data Model
ERROR: "Table 'Sales' not found"
FIX:
  excel_table(add-to-datamodel, tableName="Sales")
  excel_datamodel(create-measure, ...)

# Pattern 3: Power Query create fails - query exists
ERROR: "Query 'MyQuery' already exists"
FIX:
  excel_powerquery(update, queryName="MyQuery", ...)
```

## Data Model Workflow

Tables must be in the Data Model before DAX measures work:

1. **Worksheet Tables**: `excel_table(add-to-datamodel)`
2. **External Data**: `excel_powerquery(create, loadDestination='data-model')` → `excel_powerquery(refresh)`
3. **Create Measures**: `excel_datamodel(create-measure)`

**Power Pivot Limitations** (vs Power BI):
- NO calculated tables - use Power Query instead
- NO calculated columns via COM API - use Power Query or DAX measures
- Measures and relationships work fully

## CRITICAL: Data Model Sync After Table Changes

**Worksheet tables and Data Model are SEPARATE!** Changes don't auto-sync.

```
AFTER adding rows to worksheet table:
excel_table(append, ...)       # Worksheet updated
excel_datamodel(refresh)       # REQUIRED: Sync changes to Data Model
excel_datamodel(evaluate, ...) # Now returns updated values
```

**Skip refresh only when:** Power Query is the source (`excel_powerquery(refresh)` refreshes both)

## PivotTable Calculated Fields

For computed values in PivotTables, you have two options:

**Option 1: Calculated Field (simple, single-table)**
```
excel_pivottable_calc(CreateCalculatedField, fieldName="Revenue", formula="=Quantity*UnitPrice")
excel_pivottable_field(AddValueField, fieldName="Revenue", aggregationFunction="Sum")
```

**Option 2: DAX Measure (complex, multi-table, reusable)**
```
excel_table(add-to-datamodel, tableName="Sales")
excel_datamodel(create-measure, daxFormula="SUMX(Sales, Sales[Quantity]*Sales[UnitPrice])")
excel_pivottable(create-from-datamodel, ...)
```

**When to use DAX instead:**
- Multi-table calculations (need relationships)
- Complex logic (time intelligence, YTD, etc.)
- Reusable across multiple PivotTables

## Power Query Workflow

1. **Create Query**: `excel_powerquery(create, mCode='...')` - imports M code
2. **Load Data**: `excel_powerquery(refresh, refreshTimeoutSeconds=120)` - REQUIRED parameter
3. **Load Destinations**: `worksheet`, `data-model`, `both`, or `connection-only`

**Server Quirks:**
- `refresh` REQUIRES `refreshTimeoutSeconds` (60-600 seconds) - will fail without it
- M code is auto-formatted on create/update via powerqueryformatter.com
- `update` action auto-refreshes after updating M code
- `create` fails if query exists - use `update` instead
- `connection-only` queries NOT validated until first execution

## Chart Positioning (CRITICAL)

**NEVER place charts at default position (0,0) - they overlap data!**

```
1. excel_range(get-used-range)           # Find where data ends
2. excel_chart(create-from-range, targetRange='F2:K15')  # Place OUTSIDE data
```

**Positioning options:**
- `targetRange='F2:K15'` - cell-relative (RECOMMENDED)
- `left=400, top=200` - points (72 pts = 1 inch)

## Slicers for Visual Filtering

Create interactive filter controls for PivotTables and Tables:

```
# PivotTable slicer
excel_slicer(create-slicer, pivotTableName='Sales', fieldName='Region')

# Table slicer  
excel_slicer(create-table-slicer, tableName='SalesTable', columnName='Category')

# Set filter selection (empty array clears filter)
excel_slicer(set-slicer-selection, slicerName='RegionSlicer', selectedItems='["West","East"]')
```

## Star Schema Design

### Architecture
```
Power Query (ETL):          DAX (Business Logic):
- Load/transform data       - Calculate measures
- Create fact tables        - Time intelligence
- Create dimension tables   - Business rules
```

### Relationships
- Use `excel_datamodel_rel` to create relationships
- Pattern: Fact[ForeignKey] → Dimension[PrimaryKey]

## Named Ranges for Parameters

Use `excel_namedrange` for values Power Query can read:

1. `excel_worksheet(create)` - e.g., "_Setup"
2. `excel_range(set-values)` - parameter values
3. `excel_namedrange(create)` - named reference
4. M code: `Excel.CurrentWorkbook(){[Name = "Param_Name"]}[Content]{0}[Column1]`

## Common Patterns

### Import CSV and Build Dashboard
```
excel_file(open) → sessionId
excel_powerquery(create, loadDestination='data-model', mCode=Csv.Document(...))
excel_powerquery(refresh, refreshTimeoutSeconds=120)
excel_datamodel(create-measure, daxFormula='SUM(...)')
excel_pivottable(create-from-datamodel)
excel_file(close, save=true)
```

### Update Existing Workbook
```
excel_file(open) → sessionId
excel_powerquery(list)  # Check existing
excel_powerquery(update)  # Auto-refreshes
excel_range(get-values)   # Verify
excel_file(close, save=true)
```

## Reference Documentation

See `references/` for detailed guidance:

- @references/workflows.md - Production patterns
- @references/behavioral-rules.md - Execution guidelines
- @references/anti-patterns.md - Common mistakes
- @references/excel_pivottable.md - PivotTable operations and calculated field limitations
- @references/excel_powerquery.md - Power Query specifics
- @references/excel_datamodel.md - Data Model/DAX specifics
- @references/excel_table.md - Table operations
- @references/excel_range.md - Range operations
- @references/excel_worksheet.md - Worksheet operations
- @references/excel_chart.md - Charts, formatting, and trendlines
- @references/excel_slicer.md - Slicer operations
- @references/excel_conditionalformat.md - Conditional formatting
- @references/claude-desktop.md - Claude Desktop setup

## CLI Usage

The ExcelCLI provides the same functionality for terminal/agent automation:

```bash
# Install
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# Workflow
excelcli session open "C:\path\to\file.xlsx"  # Returns sessionId
excelcli range get-values --session <id> --sheet "Sheet1" --range "A1:D10"
excelcli session close --session <id> --save
```

Run `excelcli --help` for command discovery. Run `excelcli <command> --help` for detailed parameters.

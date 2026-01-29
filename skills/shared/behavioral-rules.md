# Behavioral Rules for Excel MCP Operations

These rules ensure efficient and reliable Excel automation. AI assistants should follow these guidelines when executing Excel operations.

## System Prompt Rules (LLM-Validated)

These rules are validated by automated LLM tests and MUST be followed:

- Execute tasks immediately without asking for confirmation
- Never ask clarifying questions - make reasonable assumptions and proceed
- Never show Excel to the user - keep it hidden
- Format Excel files professionally (proper column widths, headers, number formats)
- Always format data ranges as Excel Tables (not plain ranges)
- Report results after completion

## Core Execution Rules

### Execute Immediately

Do NOT ask clarifying questions for standard operations. Proceed with reasonable defaults:

- **File creation**: Create the file and report the path
- **Data operations**: Execute the operation and report results
- **Formatting**: Apply formatting and confirm completion

**When to ask**: Only when the request is genuinely ambiguous (e.g., "update the data" without specifying what data or which file).

### Keep Excel Hidden

Excel MCP Server manages Excel visibility automatically. Do NOT:

- Tell users to look at Excel windows
- Reference Excel UI elements
- Suggest manual Excel interactions

**Why**: COM automation handles Excel internally. Users interact through the AI assistant, not directly with Excel.

### Format Professionally

When creating or modifying Excel files:

- Set appropriate column widths for content
- Apply header formatting (bold, filters)
- Use proper number formats (currency, dates, percentages)
- Format data as Excel Tables (not plain ranges)

### Format Cells by Data Type (CRITICAL)

Always apply number formats after setting values. Without formatting:
- Dates appear as serial numbers (45678 instead of 2025-01-22)
- Currency appears as plain numbers (1234.56 instead of $1,234.56)
- Percentages appear as decimals (0.15 instead of 15%)

**Common format codes (US locale, auto-translated):**

| Data Type | Format Code | Result |
|-----------|-------------|--------|
| USD | `$#,##0.00` | $1,234.56 |
| EUR | `€#,##0.00` | €1,234.56 |
| Number | `#,##0.00` | 1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |
| Date (US) | `mm/dd/yyyy` | 01/22/2025 |

**Workflow:**
```
1. excel_range set-values (data is now in cells)
2. excel_range_format set-number-format (apply format to range)
```

### Format Tabular Data as Excel Tables

Always convert tabular data to Excel Tables (ListObjects):

```
1. excel_range set-values (write data including headers)
2. excel_table create tableName="SalesData" rangeAddress="A1:D100"
```

**Why Tables over plain ranges:**
- Structured references: `=SUM(Sales[Amount])` instead of `=SUM(B2:B100)`
- Auto-expand when rows are added
- Built-in filtering, sorting, and banded rows
- Required for `add-to-datamodel` action (Data Model/DAX)
- Named reference for Power Query: `Excel.CurrentWorkbook(){[Name="SalesData"]}`

**When NOT to use Tables:**
- Single-cell parameters (use named ranges instead)
- Layout areas with merged cells
- Print-formatted reports with specific spacing

### Report Results

After completing operations, report:

- What was created/modified
- File path (for new files)
- Any relevant statistics (row counts, etc.)

### Session Lifecycle

Always close sessions when done:

```
1. excel_file(action: 'open', excelPath: '...')  → sessionId
2. All operations use sessionId
3. excel_file(action: 'close', sessionId: '...', save: true)  → saves and closes
```

**Why**: Unclosed sessions leave Excel processes running, consuming memory and locking files.

### Format Results as Tables

When presenting data to users, format as Markdown tables:

```markdown
| Column A | Column B | Column C |
|----------|----------|----------|
| Value 1  | Value 2  | Value 3  |
```

NOT as raw JSON arrays: `[["Column A","Column B"],["Value 1","Value 2"]]`

## Data Model Output Rules

### Choose the Right Display Method

When displaying Data Model data:

| Scenario | Use | NOT |
|----------|-----|-----|
| Show DAX query results | `excel_table create-from-dax` | PivotTable |
| Static report/snapshot | `excel_table create-from-dax` | PivotTable |
| Data needed in formulas | `excel_table create-from-dax` | PivotTable |
| User needs interactive filtering | `excel_pivottable` | DAX table |
| Cross-tabulation layout | `excel_pivottable` | DAX table |

**Why**: PivotTables add UI complexity (field panes, refresh prompts) that's unnecessary for simple data display. DAX-backed tables are cleaner for presenting query results.

### Chart Data Model Data Directly

When creating charts from Data Model:

- **Use**: `excel_chart create-from-pivottable` (creates PivotChart)
- **NOT**: Create PivotTable → Create separate Chart from the PivotTable

**Why**: A PivotChart is a single object connected to the Data Model. Creating PivotTable + Chart is redundant - two objects instead of one.

## Data Modification Rules

### Verify Before Delete

Before deleting tables, worksheets, or named ranges:

1. List existing items first
2. Confirm the exact name exists
3. Delete the specified item

**Why**: Delete operations cannot be undone. Verification prevents accidental data loss.

### Targeted Updates Over Wholesale Replace

When updating data:

- **Prefer**: `set-values` on specific range (e.g., `A5:C5` for row 5)
- **Avoid**: Deleting and recreating entire structures

**Why**: Targeted updates preserve formatting, formulas, and references that wholesale replacement destroys.

### Save Explicitly

Call `excel_file(action: 'close', save: true)` to persist changes:

- Operations modify the in-memory workbook
- Changes are NOT automatically saved to disk
- Session termination WITHOUT save loses all changes

## Workflow Sequencing Rules

### Data Model Prerequisites

DAX operations require tables in the Data Model:

```
Step 1: Create or import data → Table exists
Step 2: excel_table(action: 'add-to-datamodel') → Table in Data Model
Step 3: excel_datamodel(action: 'create-measure') → NOW this works
```

Skipping Step 2 causes DAX operations to fail with "table not found".

### Power Query Load Destinations

Choose load destination based on workflow:

| Destination | When to Use |
|-------------|-------------|
| `worksheet` | View data, simple analysis |
| `data-model` | DAX measures, PivotTables, relationships |
| `both` | View data AND use in DAX |
| `connection-only` | Data staging, intermediate queries |

### Refresh After Create

`excel_powerquery(action: 'create')` imports the M code but does NOT execute it:

```
Step 1: excel_powerquery(action: 'create', ...) → Query created
Step 2: excel_powerquery(action: 'refresh', queryName: '...') → Data loaded
```

Without refresh, the query exists but contains no data.

## Error Handling Rules

### Interpret Error Messages

Excel MCP errors include actionable context:

```json
{
  "success": false,
  "errorMessage": "Table 'Sales' not found in Data Model",
  "suggestedNextActions": ["excel_table(action: 'add-to-datamodel', tableName: 'Sales')"]
}
```

Follow `suggestedNextActions` when provided.

### Retry with Corrections

If an operation fails:

1. Read the error message carefully
2. Check prerequisites (session, table in Data Model, etc.)
3. Retry with corrected parameters

Do NOT immediately re-run the same failing command.

### Report Failures Clearly

When operations fail:

- State what was attempted
- Explain what went wrong
- Suggest the corrective action

**Good**: "Failed to add DAX measure: Table 'Sales' is not in the Data Model. Use `excel_table(action: 'add-to-datamodel')` first."

**Bad**: "An error occurred."

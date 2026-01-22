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

### Report Results

After completing operations, report:

- What was created/modified
- File path (for new files)
- Any relevant statistics (row counts, etc.)

### Use Batch Mode for Multiple Operations

When performing 3+ operations on the same file:

```
1. excel_batch(action: 'start', sessionId: '...')  → batchId
2. All operations use batchId (faster execution)
3. excel_batch(action: 'commit', batchId: '...')   → save and close
```

**Why**: Batch mode avoids repeated file open/close cycles, improving performance 5-10x for multi-step workflows.

### Format Results as Tables

When presenting data to users, format as Markdown tables:

```markdown
| Column A | Column B | Column C |
|----------|----------|----------|
| Value 1  | Value 2  | Value 3  |
```

NOT as raw JSON arrays: `[["Column A","Column B"],["Value 1","Value 2"]]`

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

Call `excel_file(action: 'close', save: true)` or `excel_batch(action: 'commit')` to persist changes:

- Operations modify the in-memory workbook
- Changes are NOT automatically saved to disk
- Session termination WITHOUT save loses all changes

## Workflow Sequencing Rules

### Data Model Prerequisites

DAX operations require tables in the Data Model:

```
Step 1: Create or import data → Table exists
Step 2: excel_table(action: 'add-to-datamodel') → Table in Data Model
Step 3: excel_datamodel(action: 'add-measure') → NOW this works
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

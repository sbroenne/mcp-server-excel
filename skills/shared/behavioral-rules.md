# Behavioral Rules for Excel MCP Operations

These rules ensure efficient and reliable Excel automation. AI assistants should follow these guidelines when executing Excel operations.

## Core Execution Rules

### Discover, then stay within scope

Use session, worksheet, and table listings to discover existing state. Match
the intended workbook; do not choose an unrelated session or invent a path.
Ask when the requested target or a destructive change remains unclear.
Reading a workbook does not require writing, formatting, Tables, or PivotTables.

Excel is hidden by default. Show it when requested, without imposing a visibility
menu. Confirm before closing a visible window unless already authorized.

### Format Professionally

When formatting is part of the task:

- Set appropriate column widths for content
- Apply header formatting (bold, filters)
- Use proper number formats (currency, dates, percentages) with `range set-number-format`
- Auto-fit variable-width data with `range_format auto-fit-columns` or `range_format auto-fit-rows`
- Create Excel Tables when requested or required for the intended workflow
- When the same visual styling applies to multiple disjoint ranges on one sheet, use `range_format format-ranges`

**Tool split to remember:**
- `range` owns number display formats such as dates, currency, percentages, and text display
- `range_format` owns visual styling, validation, auto-fit, and explicit width/height changes

**Use `set-style` for semantic status labels and document structure:**
- `Good` / `Bad` / `Neutral` — colour-coded status cells (green/red/yellow fills, theme-aware)
- `Heading 1` / `Heading 2` / `Title` — document hierarchy
- `Normal` — reset all formatting

**Use `format-range` for visual layout (header rows, custom colours) — ALL properties in ONE call:**
- `set-style('Heading 1')` does NOT apply a fill colour; if you want a coloured header row use `format-range`
- Pass `bold`, `fill_color`, `font_color`, and alignment together in a single call — do not call `format-range` multiple times for the same range
- If the same formatting payload repeats across multiple non-contiguous ranges, prefer one `format-ranges` call over repeated `format-range` calls

**Apply each formatting operation once** — do not reapply the same properties to the same range unless a later step explicitly changes them.

### Format Cells by Data Type (CRITICAL)

Preserve existing formats unless the task calls for changing them. New numeric
data may need number formats:
- Dates appear as serial numbers (45678 instead of 2025-01-22)
- Currency appears as plain numbers (1234.56 instead of $1,234.56)
- Percentages appear as decimals (0.15 instead of 15%)

**Common format codes (US locale, auto-translated):**

| Data Type | Format Code | Result (en-US) |
|-----------|-------------|----------------|
| USD | `$#,##0.00` | $1,234.56 |
| EUR | `€#,##0.00` | €1,234.56 |
| Number | `#,##0.00` | 1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |
| Date (US) | `mm/dd/yyyy` | 01/22/2025 |

**Rendered output is locale-dependent.** The `Result` column assumes en-US regional settings. Excel
interprets `,` and `.` in a format code according to the user's locale, so `$#,##0.00` displays as
`$1,234.56` on en-US but `$1.234,56` on de-DE — same code, different separators. Always write the US
form (it is auto-translated), never promise a literal rendering, and never "correct" a format code
because a screenshot shows swapped separators.

**Workflow:**
```
1. range set-values (data is now in cells)
2. range set-number-format (apply format to range)
3. range_format auto-fit-columns (widen columns to fit — formatted dates and
   long numbers render as ##### at the default column width)
```

### Format Tabular Data as Excel Tables

When an Excel Table (ListObject) is requested or needed:

```
1. range set-values (write data including headers)
2. table(action: 'create', table_name: 'SalesData', range_address: 'A1:D100')
```

**Why Tables over plain ranges:**
- Structured references: `=SUM(Sales[Amount])` instead of `=SUM(B2:B100)`
- Auto-expand when rows are added
- Built-in filtering, sorting, and banded rows
- Required for `add-to-data-model` action (Data Model/DAX)
- Named reference for Power Query: `Excel.CurrentWorkbook(){[Name="SalesData"]}`

**When NOT to use Tables:**
- Single-cell parameters (use named ranges instead)
- Layout areas with merged cells
- Print-formatted reports with specific spacing

**Named range listing:** `namedrange list` returns visible user-defined names. Hidden/internal Excel names, including Power Query `ExternalData_*` and AutoFilter names, are omitted before value inspection. Large named ranges return metadata without a value preview; use `namedrange read` or `range get-values` when the actual value is needed.

### Report Results

After completing operations, report:

- What was created/modified
- File path (for new files)
- Any relevant statistics (row counts, etc.)

### Session Lifecycle

Use `file(action: 'test')` or `excelcli -q session test <path>` before opening
when access or information protection is uncertain. The shared result reports
`canOpen`, `isIrmProtected`, `willOpenReadOnly`, and `requiresVisibleSession`.
For ordinary workbooks, this briefly opens the file read-only in Excel and
closes it without saving; use the timeout option for slow validation opens.
IRM/AIP files report `canOpen:false` until the required interactive Excel
authentication occurs; open them with a visible session.

Close when authorized, work is finished, and `file list` reports `canClose: true`.
Keep the workbook open when requested. MCP example:

```
1. file(action: 'open', path: '...')  → capture response.session_id as sessionId
2. workbook(action: 'get-info', session_id: sessionId)
3. file(action: 'close', session_id: sessionId, save: true)  → saves and closes
```

Pass that same value as `session_id` on every session-based MCP follow-up.
`sessionId` above is a local variable, not an MCP argument name. For `file(list)`,
copy the matching entry's `sessionId` value into `session_id`; never guess a session.
The MCP server has a defensive bridge-compatibility fallback for a top-level
`sessionId`, but agents must continue to send the canonical `session_id`.
Compatibility use is recorded with a privacy-safe warning and telemetry signal.

CLI commands instead return `sessionId` and accept `--session`:

```powershell
$session = excelcli -q session open C:\path\file.xlsx | ConvertFrom-Json
$sessionId = $session.sessionId
excelcli -q workbook get-info --session $sessionId
excelcli -q session close --session $sessionId --save
```

CLI and MCP sessions are separate; IDs cannot be transferred between them.

**Why**: Unclosed sessions leave Excel processes running, consuming memory and locking files.

## Data Model Output Rules

### Choose the Right Display Method

When displaying Data Model data:

| Scenario | Use | NOT |
|----------|-----|-----|
| Show DAX query results | `table create-from-dax` | PivotTable |
| Static report/snapshot | `table create-from-dax` | PivotTable |
| Data needed in formulas | `table create-from-dax` | PivotTable |
| User needs interactive filtering | `pivottable` | DAX table |
| Cross-tabulation layout | `pivottable` | DAX table |

**Why**: PivotTables add UI complexity (field panes, refresh prompts) that's unnecessary for simple data display. DAX-backed tables are cleaner for presenting query results.

### Chart Data Model Data Directly

When creating charts from Data Model:

- **Use**: `chart create-from-pivottable` (creates and verifies a live PivotChart)
- **NOT**: Create a regular chart from the PivotTable's displayed cell range

**Why**: The verified PivotLayout link keeps field changes and refreshes live.

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

Call `file(action: 'close', save: true)` to persist changes:

- Operations modify the in-memory workbook
- Explicit close defaults to discarding unsaved edits
- Normal Service shutdown attempts to save remaining sessions before disposal
- Bare batch disposal, crashes, timeouts, and forced cleanup do not guarantee saving
- Cancellation is not undo; inspect the session list before continuing or retrying

## Workflow Sequencing Rules

### Data Model Prerequisites

DAX operations require tables in the Data Model:

```
Step 1: Create or import data → Table exists
Step 2: table(action: 'add-to-data-model') → Table in Data Model
Step 3: datamodel(action: 'create-measure') → NOW this works
```

Skipping Step 2 causes DAX operations to fail with "table not found".

### Power Query Load Destinations

Choose load destination based on workflow:

| Destination | When to Use |
|-------------|-------------|
| `worksheet` / `load-to-table` | View data, simple analysis |
| `data-model` / `load-to-data-model` | DAX measures, PivotTables, relationships |
| `both` / `load-to-both` | View data AND use in DAX |
| `connection-only` | Data staging, intermediate queries |

Values are case-insensitive. Unknown enum values and parameters that do not
belong to the selected action are rejected instead of being defaulted or ignored.

### Canonical Public Inputs

- Public timeouts are integer seconds: MCP uses `timeout_seconds`, CLI uses
  `--timeout`, and batch JSON uses `timeout`. Do not send TimeSpan strings or
  numeric strings.
- Session open/create accepts 10-3600 seconds. Its operation timeout controls
  workbook startup and operations that do not provide a dedicated data-operation
  timeout.
- Power Query refresh/refresh-all accepts 0-2147483; omitted or `0` uses the
  30-minute data-operation default. Connection, Data Model, PivotTable, and VBA
  timeouts accept 1-2147483.
- Power Query refresh/refresh-all use their data-operation timeout instead of
  layering the session operation timeout. `load-to` has no caller timeout and
  uses the fixed 30-minute data-operation timeout; create/update/evaluate use the
  session operation timeout.
- For required generated inline/file pairs, supply exactly one form. Optional
  pairs may omit both, but inline and file forms are always mutually exclusive.
  Batch aliases are
  `mCodeFile`, `vbaCodeFile`, `daxFormulaFile`, `daxQueryFile`, `dmvQueryFile`,
  `schemaFile`, and `xmlDataFile`; MCP uses snake_case and CLI uses kebab-case.
  The file must exist and be readable.

### Create, Load, and Refresh

`powerquery create` stores M code and loads its selected destination
(`worksheet` by default). Choose `connection-only` to store without execution.
Use `load-to` to change destinations, and `refresh` when loaded data needs updating.
Update refreshes by default unless `refresh: false` is supplied.

```
powerquery(action: 'create', load_destination: 'data-model', ...)
datamodel(action: 'list-tables', ...)  // Data is already loaded
```

## Error Handling Rules

### Interpret Error Messages

Excel MCP errors include actionable context:

```json
{
  "success": false,
  "errorMessage": "Table 'Sales' not found in Data Model",
  "suggestedNextActions": ["table(action: 'add-to-data-model', table_name: 'Sales')"]
}
```

Follow `suggestedNextActions` when provided.

Use the structured `errorCategory` when it is present. For `InvalidInput`,
`NotFound`, or `Conflict`, correct the named input or workbook state before
retrying. For `SessionNotFound`, timeout, or cancellation, inspect `file list`
and reuse a surviving matching session or reopen the known workbook when
necessary. Inspect affected data before continuing: a failed operation may
have partly applied, and reopening cannot recover unsaved edits.

`Prerequisite` means required workbook data or a feature is missing, such as
tables in the Data Model. `DependencyUnavailable` means an external component
is missing, such as the MSOLAP provider. `Permissions` can indicate blocked VBA
project access; do not change security settings automatically. Running an
existing macro does not itself require VBA project access. Unknown Excel
errors remain unknown: do not assume a generic COM failure is a bad query or
a trust problem.

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

**Good**: "Failed to add DAX measure: Table 'Sales' is not in the Data Model. Use `table(action: 'add-to-data-model')` first."

**Bad**: "An error occurred."

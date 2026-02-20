# Anti-Patterns to Avoid

These patterns cause data loss, poor performance, or user frustration. Avoid them.

## Redundant Formatting Anti-Pattern

### The Problem

Applying the same formatting to the same range more than once in a workflow:

```
WRONG: Applying bold repeatedly

range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true)
// ... other operations ...
range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true, fillColor: '#4472C4')
// Bold was already applied - the second call re-applies it unnecessarily
```

Also wrong: calling `format-range` separately for each property instead of combining:

```
WRONG: Separate calls for each property

range_format(action: 'format-range', rangeAddress: 'A1:D1', bold: true)
range_format(action: 'format-range', rangeAddress: 'A1:D1', fillColor: '#4472C4')
range_format(action: 'format-range', rangeAddress: 'A1:D1', fontColor: '#FFFFFF')
range_format(action: 'format-range', rangeAddress: 'A1:D1', horizontalAlignment: 'center')
```

### The Solution

Apply all formatting properties for a range in **one** `format-range` call:

```
CORRECT: One call per range

range_format(action: 'format-range', rangeAddress: 'A1:D1',
    bold: true, fillColor: '#4472C4', fontColor: '#FFFFFF', horizontalAlignment: 'center')
```

Apply each formatting operation **once**. If a subsequent step explicitly changes a property (e.g., "now make the title red"), apply it again — otherwise don't.

### When Multiple Calls ARE Appropriate

- Applying different formatting to different ranges
- A later step explicitly overrides a previously set property
- Applying a style (`set-style`) on top of individual properties (different actions)



### The Problem

Deleting entire structures to make small changes:

```
WRONG: User wants to update cell B5

table(action: 'delete', tableName: 'SalesData')
range(action: 'set-values', values: [[entire dataset with B5 fixed]])
table(action: 'create', tableName: 'SalesData', ...)
```

This destroys:
- Cell formatting
- Conditional formatting rules
- Data validation
- Named ranges pointing to the table
- PivotTable connections
- DAX measures referencing the table

### The Solution

Use targeted modifications:

```
CORRECT: Update only the changed cell

range(action: 'set-values', rangeAddress: 'B5', values: [[newValue]])
```

### When Rebuild IS Appropriate

- Fundamentally restructuring data (different columns)
- Converting between table types
- User explicitly requests replacement

## Discovery Loop Anti-Pattern

### The Problem

Repeating `file(list)`, `worksheet(list)`, or `table(list)` multiple times without taking action:

```
WRONG: Looping on discovery after an error

worksheet(action: 'list')           → gets sheet list
worksheet(action: 'list')           → gets same sheet list again
file(action: 'list')                → gets session list
worksheet(action: 'list')           → gets same sheet list again
... (dozens of repetitions)
```

This burns tokens, costs money, and never completes the task.

### The Solution

If you already have a sessionId, use it. Do not rediscover:

```
CORRECT: Use the sessionId you already have

Error: "session expired"
→ file(action: 'open', path: original_path)  ← Re-open once, get new sessionId
→ Continue with the new sessionId immediately
```

### The Rule

- **Max 2 retries** for any session or file operation
- After 2 failures: stop retrying, report the error, end your response
- **Never call `list`, `worksheet(list)`, or `table(list)` more than twice in a row** without doing something with the result

## Confirmation Loop Anti-Pattern

### The Problem

Asking for confirmation on every operation:

```
WRONG:

User: "Create a sales report"
AI: "Would you like me to create a new Excel file for the sales report?"
User: "Yes"
AI: "What would you like to name the file?"
User: "sales_report.xlsx"
AI: "Should I create it in your Documents folder?"
User: "Yes"
AI: "The file has been created. Would you like me to add headers?"
... (10 more questions)
```

### The Solution

Execute with reasonable defaults, report results:

```
CORRECT:

User: "Create a sales report"
AI: "Created sales report at C:\Users\You\Documents\sales_report.xlsx with the following structure:
- Sheet 'Summary' with headers: Date, Product, Region, Sales
- Ready for data entry

What data would you like to add?"
```

### When to Ask

- Genuinely ambiguous requests
- Destructive operations on existing data
- User explicitly asked for options

## Wrong Cell Update Anti-Pattern

### The Problem

Reading entire range, modifying in memory, writing entire range back:

```
WRONG: Update one cell by rewriting thousands

data = range(action: 'get-values', rangeAddress: 'A1:Z1000')
data[4][1] = "new value"  // Modify row 5, column B
range(action: 'set-values', rangeAddress: 'A1', values: data)
```

This:
- Transfers megabytes unnecessarily
- Risks data corruption if interrupted
- Destroys formulas (values only, not formulas)
- Loses cell formatting

### The Solution

Write only the changed cells:

```
CORRECT: Direct cell update

range(action: 'set-values', rangeAddress: 'B5', values: [["new value"]])
```

## Session Leak Anti-Pattern

### The Problem

Opening files without closing them:

```
WRONG: Session accumulation

file(action: 'open', filePath: 'file1.xlsx')  // Session 1
file(action: 'open', filePath: 'file2.xlsx')  // Session 2
file(action: 'open', filePath: 'file3.xlsx')  // Session 3
// ... never closed
```

Results:
- Excel processes accumulate
- Memory usage grows
- File locks prevent other access
- System becomes unresponsive

### The Solution

Always close sessions:

```
CORRECT: Proper lifecycle

session1 = file(action: 'open', path: 'file1.xlsx')
// ... work with file1 ...
file(action: 'close', sessionId: session1, save: true)

session2 = file(action: 'open', path: 'file2.xlsx')
// ... work with file2 ...
file(action: 'close', sessionId: session2, save: true)
```

## Ignoring Error Context Anti-Pattern

### The Problem

Retrying failed operations without reading the error:

```
WRONG: Blind retry

datamodel(action: 'create-measure', ...) → Error: Table not in Data Model
datamodel(action: 'create-measure', ...) → Error: Table not in Data Model
datamodel(action: 'create-measure', ...) → Error: Table not in Data Model
```

### The Solution

Read and act on error context:

```
CORRECT: Error-driven correction

datamodel(action: 'create-measure', ...) 
→ Error: Table 'Sales' not in Data Model
→ Suggested: table(action: 'add-to-data-model', tableName: 'Sales')

table(action: 'add-to-data-model', tableName: 'Sales')  // Fix prerequisite
datamodel(action: 'create-measure', ...)  // Now succeeds
```

## Number Format Locale Anti-Pattern

### The Problem

Using locale-specific format codes:

```
WRONG: German/European format

range(action: 'set-number-format', formatCode: '#.##0,00')  // German
range(action: 'set-number-format', formatCode: '# ##0,00')  // French
```

### The Solution

Always use US format codes (Excel translates automatically):

```
CORRECT: US format codes (universal)

range(action: 'set-number-format', formatCode: '#,##0.00')
```

Excel displays the result in the user's locale setting, but the API requires US format input.

## Load Destination Mismatch Anti-Pattern

### The Problem

Wrong load destination for the workflow:

```
WRONG: Loading to worksheet when DAX is needed

powerquery(action: 'create', loadDestination: 'worksheet', ...)
datamodel(action: 'create-measure', ...)  // FAILS: table not in Data Model
```

### The Solution

Match load destination to workflow:

```
CORRECT: Load to Data Model for DAX workflows

powerquery(action: 'create', loadDestination: 'data-model', ...)
powerquery(action: 'refresh', ...)
datamodel(action: 'create-measure', ...)  // Works
```

| Workflow Goal | Load Destination |
|---------------|------------------|
| View data in cells | `worksheet` |
| Use in DAX/PivotTables | `data-model` |
| Both viewing and DAX | `both` |
| Intermediate staging | `connection-only` |

## Skipping Power Query Evaluate Anti-Pattern

### The Problem

Creating or updating Power Query queries without testing M code first:

```
WRONG: Creating permanent query with untested M code

powerquery(action: 'create', mCode: '...', ...)
// M code has syntax error → COM exception with cryptic message
// Now workbook is polluted with broken query
```

This causes:
- Broken queries persisted in workbook
- Cryptic COM exceptions instead of helpful M error messages
- Need manual Excel cleanup to remove broken queries
- Wasted time debugging in wrong layer

### The Solution

Always evaluate M code BEFORE creating permanent queries:

```
CORRECT: Test-first development workflow

// Step 1: Test M code without persisting
powerquery(action: 'evaluate', mCode: '...')
// → Returns actual data preview with columns and rows
// → Better error messages if M code has issues

// Step 2: Create permanent query with validated code
powerquery(action: 'create', mCode: '...', ...)

// Step 3: Load data to destination
powerquery(action: 'refresh', ...)
```

**Benefits:**
- Catch syntax errors and missing sources BEFORE persisting
- See actual data preview (columns, sample rows)
- Better error messages than COM exceptions
- No cleanup needed - temporary objects auto-deleted
- Like a REPL for M code

### When Evaluate IS Optional

- Trivial literal tables: `#table({"Column1"}, {{123}})`
- M code already validated in previous evaluate call
- Copying known-working query from another workbook

### When to Retry With Evaluate

If create/update fails with COM error, use evaluate to get detailed Power Query error message:

```
powerquery(action: 'create', ...)  // → COM exception
powerquery(action: 'evaluate', mCode: '...')  // → Detailed M error
// Fix M code based on error
powerquery(action: 'create', ...)  // → Success
```

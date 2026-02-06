# excel_powerquery - Server Quirks

## RECOMMENDED DEVELOPMENT WORKFLOW (ALWAYS USE THIS)

**Test BEFORE persisting - avoid polluting workbooks with broken queries:**

```
Step 1: evaluate → Test M code, verify results (catches syntax errors, missing sources)
Step 2: create/update → Store VALIDATED query in workbook
Step 3: refresh/load-to → Load data to destination (worksheet/data-model)
```

**Why this workflow:**
- `evaluate` executes M code WITHOUT creating permanent query (test-then-commit)
- Returns actual data preview with columns and rows in JSON
- Better error messages than COM exceptions from create/update
- No cleanup needed - temporary objects auto-deleted
- Skip evaluate only for trivial literal tables (`#table` with hardcoded values)

**IF CREATE/UPDATE FAILS**: Use `evaluate` to get detailed Power Query error message, fix code, retry.

**Additional evaluate use cases:**
- Execute one-off queries without creating permanent queries
- Ad-hoc data exploration or debugging M code transformations
- Quick testing during development (like REPL for M code)

---

**Automatic M-Code Formatting**:

- Create and Update operations automatically format M code using powerqueryformatter.com API
- Formatting adds ~100-500ms network latency per call
- Graceful fallback: returns original M code if formatting service unavailable
- Read operations (List, View) return M code as stored (no formatting on read)
- Formatting improves readability with proper indentation, spacing, and line breaks

**Data Model workflow**:

Power Query can load data to different destinations:
- `worksheet` (default): Creates an Excel Table on a worksheet
- `data-model`: Loads directly to Power Pivot for DAX analysis
- `both`: Loads to worksheet AND Power Pivot
- `connection-only`: Imports query definition without loading data

To create DAX measures on Power Query data:
1. Use excel_powerquery create/load-to with `loadDestination='data-model'`
2. Then use excel_datamodel to create DAX measures

Alternative path (for existing worksheet tables):
1. Use excel_table with `add-to-data-model` action
2. Then use excel_datamodel to create DAX measures

**Action disambiguation**:

- **evaluate**: **CRITICAL - USE THIS FIRST** - Execute M code directly, return results WITHOUT creating a permanent query (test before create/update!)
- create: Import NEW query using inline `mCode` (FAILS if query already exists - use update instead)
- update: Update EXISTING query M code + refresh data (use this if query exists)
- rename: Change query name (requires both `queryName` and `newName` parameters)
- load-to: Loads to worksheet or data model or both (not just config change) - CHECKS for sheet conflicts
- unload: Removes data from ALL destinations (worksheet AND Data Model) - keeps query definition
- delete: Completely removes query AND all associated data (worksheet, Data Model connections)

**Rename behavior**:

- Names are trimmed and compared case-insensitively for uniqueness
- Renaming "Query1" to "query1" is allowed (case-only change, no conflict)
- Renaming "Query1" to " Query1 " is a no-op (trimmed names match)
- No-op (same normalized name) → success with `oldName` = `newName`
- Conflict with existing query → error with `errorMessage`
- M code content is unchanged - only the name changes
- No auto-save: workbook must be saved separately to persist the rename

**When to use create vs update**:

- Query doesn't exist? → Use create
- Query already exists? → Use update (create will error "already exists")
- Not sure? → Check with list action first, then use update if exists or create if new
- **ALWAYS evaluate M code FIRST** to catch errors before persisting

**List action and IsConnectionOnly**:

- `IsConnectionOnly=true` means query has NO data destination (not in worksheet, not in Data Model)
- `IsConnectionOnly=false` means query loads data SOMEWHERE (worksheet OR Data Model OR both)
- A query loaded ONLY to Data Model is NOT connection-only

**Inline M code**:

- Provide raw M code directly via `mCode`
- Keep `.pq` files only for GIT workflows

**Create/LoadTo with existing sheets**:

- Use `targetCellAddress` to place the table on an existing worksheet without deleting other content
- Applies to BOTH create and load-to
- If the worksheet already has data and you omit `targetCellAddress`, the tool returns guidance telling you to provide one
- Existing tables are refreshed in-place; specifying a different `targetCellAddress` requires unload + reload
- Worksheets that exist but are empty behave like new sheets (default destination = A1)

**Common mistakes**:

- **WARNING: Skipping evaluate** → Create/update with untested M code (ERROR: pollutes workbook with broken queries)
- Using create on existing query → ERROR "Query 'X' already exists" (should use update)
- Using update on new query → ERROR "Query 'X' not found" (should use create)
- Calling LoadTo without checking if sheet exists (will error if sheet exists)
- Assuming unload only removes worksheet data → Also removes Data Model connections
- Calling rename without trimming newName → Server trims automatically, " Query " becomes "Query"
- Renaming to conflicting name → Check list first if unsure about existing names

**Server-specific quirks**:

- Validation = execution: M code only validated when data loads/refreshes
- connection-only queries: NOT validated until first execution
- refresh with loadDestination: Applies load config + refreshes (2-in-1)
- Single cell returns [[value]] not scalar
- refresh action REQUIRES `refreshTimeoutSeconds` between 60-600 seconds (1-10 minutes). If refresh needs more than 10 minutes, ask the user to run it manually in Excel—the server refuses longer windows and will not pick a default for you.
- load-to has a 5-minute guard. If Excel is blocked by privacy dialogs/credentials, you'll get `SuggestedNextActions` instead of a hang—surface them to the user before retrying.

**Data Model connection cleanup**:

- Unload removes BOTH worksheet ListObjects AND Data Model connections
- Delete removes query, worksheet ListObjects, AND Data Model connections
- Connection naming pattern: "Query - {queryName}" or "Query - {queryName} - suffix"

## M Code - Server-Specific Notes

### Column/Field Name Quoting (CRITICAL)

M code requires special syntax for identifiers containing hyphens, spaces, or special characters:

| Column Name | Syntax | Notes |
|-------------|--------|-------|
| `Amount` | `[Amount]` | Simple names work without quotes |
| `Non-Recurring` | `[#"Non-Recurring"]` | **Hyphen requires `#"..."` quoting** |
| `List Price (USD)` | `[#"List Price (USD)"]` | Spaces/parens require quoting |
| `Service Level 1` | `[#"Service Level 1"]` | Spaces require quoting |

**Common mistake:** `[Non-Recurring]` parses as `[Non] - [Recurring]` (subtraction!) and fails with cryptic "The name 'X' wasn't recognized" errors.

**Rule:** If a column name contains anything other than letters, numbers, and underscores, use `[#"Column Name"]` syntax.

### Reading Named Ranges (parameters)

```m
Excel.CurrentWorkbook(){[Name = "Param_Name"]}[Content]{0}[Column1]
```

### Query Chaining

Reference other queries by name directly: `Source = OtherQueryName`

### Source Control Pattern

1. Store M code in `.pq` files
2. `excel_powerquery create` or `update` with inline `mCode`
3. `refresh` to validate
4. File name MUST match query name

Query naming: File name MUST match Excel query name exactly.

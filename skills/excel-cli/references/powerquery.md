> **CLI syntax note:** This shared domain guide may use MCP-style `tool(action: ...)` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live `--help`; notably, MCP `file` open/close maps to CLI `session` open/close, and MCP `worksheet` maps to CLI `sheet`.

# powerquery - Server Quirks

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
- Temporary queries, sheets, tables, and mashup connections are deleted by exact
  `Location`; cleanup failures return an error instead of success
- Skip evaluate only for trivial literal tables (`#table` with hardcoded values)

**IF CREATE/UPDATE FAILS**: Use `evaluate` to get detailed Power Query error message, fix code, retry.

**Additional evaluate use cases:**
- Execute one-off queries without creating permanent queries
- Ad-hoc data exploration or debugging M code transformations
- Quick testing during development (like REPL for M code)

---

**M-Code Formatting and reads**:

- Create and Update preserve M code exactly by default and do not call remote services
- Set `format_m_code=true` only with explicit user consent; it sends M code to powerqueryformatter.com
- Remote formatting adds ~100-500ms network latency per call
- Graceful fallback: saves original M code if the formatting service is unavailable
- `list` returns compact metadata, exact load mode, character count, and at most
  80 characters of `formulaPreview`; it never returns full M code
- Use `view` to retrieve one query's complete M code and exact load mode
- Read operations return stored text without formatting

**Data Model workflow**:

Power Query can load data to different destinations:
- `worksheet` / `load-to-table` (default): Creates an Excel Table on a worksheet
- `data-model` / `load-to-data-model`: Loads directly to Power Pivot for DAX analysis
- `both` / `load-to-both`: Loads to worksheet AND Power Pivot
- `connection-only`: Imports query definition without loading data

Destination values are case-insensitive. Unknown values fail before the query is
changed; they never fall back to connection-only or another enum default.

To create DAX measures on Power Query data:
1. Use powerquery create/load-to with `load_destination='data-model'`
2. Then use datamodel to create DAX measures

Alternative path (for existing worksheet tables):
1. Use table with `add-to-data-model` action
2. Then use datamodel to create DAX measures

**Action disambiguation**:

- **evaluate**: **CRITICAL - USE THIS FIRST** - Execute M code directly, return results WITHOUT creating a permanent query (test before create/update!)
- create: Import NEW query using inline `m_code` (FAILS if query already exists - use update instead)
- update: Update EXISTING query M code + refresh data (use this if query exists)
- rename: Change query name (requires both `old_name` and `new_name`)
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

**List/view load state**:

- `list`, `view`, and `get-load-config` use the same exact-identity detector
- `loadMode` distinguishes `connection-only`, `load-to-table`,
  `load-to-data-model`, and `load-to-both`
- `IsConnectionOnly=true` means query has NO data destination (not in worksheet, not in Data Model)
- `IsConnectionOnly=false` means query loads data SOMEWHERE (worksheet OR Data Model OR both)
- A query loaded ONLY to Data Model is NOT connection-only
- If Excel cannot inspect a query, `list` fails explicitly instead of silently
  omitting that query

**Inline M code**:

- Provide raw M code directly via `m_code`
- Keep `.pq` files only for GIT workflows

**Create/LoadTo with existing sheets**:

- Use `target_cell_address` to place the table on an existing worksheet without deleting other content
- Applies to BOTH create and load-to
- If the worksheet already has data and you omit `target_cell_address`, the tool returns guidance telling you to provide one
- Existing tables are refreshed in-place; specifying a different `target_cell_address` requires unload + reload
- Worksheets that exist but are empty behave like new sheets (default destination = A1)

**Common mistakes**:

- **WARNING: Skipping evaluate** → Create/update with untested M code (ERROR: pollutes workbook with broken queries)
- Using create on existing query → ERROR "Query 'X' already exists" (should use update)
- Using update on new query → ERROR "Query 'X' not found" (should use create)
- Calling LoadTo without checking if sheet exists (will error if sheet exists)
- Assuming unload only removes worksheet data → Also removes Data Model connections
- Calling rename without trimming `new_name` → Server trims automatically, " Query " becomes "Query"
- Renaming to conflicting name → Check list first if unsure about existing names
- Passing an option from another action (for example `m_code` on delete or `timeout_seconds` on load-to) → ERROR; category-wide schemas expose the union of options, but each action validates its own subset

**Server-specific quirks**:

- Validation = execution: M code only validated when data loads/refreshes
- connection-only queries: NOT validated until first execution
- load-to with `load_destination`: Applies load config + refreshes (2-in-1)
- Single cell returns [[value]] not scalar
- Public timeout inputs are integer seconds. Refresh/refresh-all accepts 0-2147483 and defaults to the 30-minute data-operation timeout when `timeout_seconds`/`--timeout` is 0 or omitted. For quick queries use a smaller value (e.g., 60-120 seconds).
- load-to has no caller timeout parameter and uses the fixed 30-minute data-operation timeout; passing `timeout` to load-to is rejected instead of ignored.
- Refresh/refresh-all data-operation timeouts replace the session operation wait rather than layering with it. The session timeout controls startup and operations without a dedicated data timeout, including create/update/evaluate. Session open/create accepts 10-3600 seconds and defaults to 120.

**Data Model connection cleanup**:

- Unload removes BOTH worksheet ListObjects AND Data Model connections
- Delete removes query, worksheet ListObjects, AND Data Model connections
- Ownership uses the exact case-insensitive mashup `Location` value, not connection
  display names; queries such as `A` and `AA` remain isolated

## M Code - Server-Specific Notes

> For full M code language syntax, see [m-code-syntax reference](m-code-syntax.md).

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
2. `powerquery create` or `update` with inline `m_code`
3. `refresh` to validate
4. File name MUST match query name

Query naming: File name MUST match Excel query name exactly.

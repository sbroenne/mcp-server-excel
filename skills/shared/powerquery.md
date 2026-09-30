# powerquery - Server Quirks

## Test New or Changed Queries Within the Request

For authorized query development, prefer testing new logic before storing it:

```
Step 1: evaluate → Test M code, verify results (catches syntax errors, missing sources)
Step 2: create/update → Store the validated query; create also loads its chosen destination
Step 3: load-to if created connection-only; refresh when loaded data needs updating
```

**Why this workflow:**
- `evaluate` executes M code without keeping a permanent query
- Returns actual data preview with columns and rows in JSON
- Better error messages than COM exceptions from create/update
- Temporary queries, sheets, tables, and mashup connections are deleted by exact
  `Location`; cleanup failures return an error instead of success
- Skip redundant evaluation for trivial literal tables or already-validated code
  with unchanged sources and dependencies

Evaluation is not a read-only operation: it creates a temporary query, sheet,
Table, and connection in the current workbook, executes M code, then removes
those objects. Execution may contact external sources. An audit does not
authorize evaluation, refresh, or repairs merely because the resulting objects
are temporary. Use stored definitions, metadata, and existing loaded values for
inspection; follow the shared [permission policy](behavioral-rules.md#intent-and-permission)
when a state-changing check is necessary.

## Recovering a failed create

Creation adds the query before loading. A failed load can leave the query and
load objects behind. Inspect `list`, `view`, and `get-load-config` first. Evaluate
corrected code, then **update if the query survived**; create only if it is absent.
Do not delete surviving objects blindly.

For an existing query `SalesQuery`, with corrected code in a known readable
`query.m` file and the current session already captured:

```mcp
powerquery(action: 'evaluate', session_id: sessionId, m_code_file: 'query.m')
powerquery(action: 'update', session_id: sessionId, query_name: 'SalesQuery', m_code_file: 'query.m', refresh: false)
powerquery(action: 'get-load-config', session_id: sessionId, query_name: 'SalesQuery')
```

```cli
excelcli -q powerquery evaluate --session $sessionId --m-code-file query.m
excelcli -q powerquery update --session $sessionId --query-name SalesQuery --m-code-file query.m --refresh false
excelcli -q powerquery get-load-config --session $sessionId --query-name SalesQuery
```

Check each result. Refresh a surviving intended load, or use `load-to` for the
required destination after inspecting sheet content. A successful evaluation does
not prove that loading onto a particular sheet will succeed.

**M-Code Formatting and reads**:

- Create and Update preserve M code exactly by default and do not call remote services
- Enable remote formatting only with explicit user consent; it sends M code to powerqueryformatter.com
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
1. Create/load the query with a Data Model destination
2. Then use datamodel to create DAX measures

Alternative path (for existing worksheet tables):
1. Use table with `add-to-data-model` action
2. Then use datamodel to create DAX measures

**Action disambiguation**:

- evaluate: Execute M code and return results without keeping a permanent query
- create: Store a new query and load its selected destination (fails if it already exists; use update instead)
- update: Update an existing query; refresh defaults to true and can be disabled
- rename: Change query name (requires the old and new names)
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
- Prefer evaluate for new or changed code to catch errors before persisting

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

- Supply raw M code inline, or use the file input for a readable file containing
  longer M code. Do not supply both.

**Create/LoadTo with existing sheets**:

- Supply a target cell address to place the Table on an existing worksheet without deleting other content
- Applies to BOTH create and load-to
- If the worksheet already has data and no target cell address is supplied, the
  tool returns guidance telling you to provide one
- Existing Tables are refreshed in place. Moving a load requires unload and
  reload, which removes existing destinations and must be within the authorized
  scope; a requested new position is not permission to discard unrelated data.
- Worksheets that exist but are empty behave like new sheets (default destination = A1)

**Common mistakes**:

- Persisting untested code can leave a broken query; prefer evaluate for new logic
- Using create on existing query → ERROR "Query 'X' already exists" (should use update)
- Using update on new query → ERROR "Query 'X' not found" (should use create)
- Loading onto populated sheets without a suitable target cell address or into overlapping content
- Assuming unload only removes worksheet data → Also removes Data Model connections
- Assuming rename preserves outer whitespace; the server trims " Query " to "Query"
- Renaming to conflicting name → Check list first if unsure about existing names
- Passing an option from another action, such as M code on delete or a caller
  timeout on load-to. Category-wide schemas expose the union of options, but
  each action validates its own subset.

**Server-specific quirks**:

- Validation = execution: M code only validated when data loads/refreshes
- connection-only queries: NOT validated until first execution
- load-to: Applies the selected load configuration and refreshes
- Single cell returns [[value]] not scalar
- Public timeout inputs are integer seconds. Refresh/refresh-all accepts
  0-2147483; zero or omission uses the 30-minute data-operation default.
- load-to has no caller timeout parameter and uses the fixed 30-minute data-operation timeout; passing `timeout` to load-to is rejected instead of ignored.
- Refresh/refresh-all data-operation timeouts replace the session operation wait rather than layering with it. The session timeout controls startup and operations without a dedicated data timeout, including create/update/evaluate. Session open/create accepts 10-3600 seconds and defaults to 120.

**Data Model connection cleanup**:

- Unload removes BOTH worksheet ListObjects AND Data Model connections
- Delete removes query, worksheet ListObjects, AND Data Model connections
- Ownership uses the exact case-insensitive mashup `Location` value, not connection
  display names; queries such as `A` and `AA` remain isolated

## M Code Files

See the [M-syntax reference](m-code-syntax.md) for quoted identifiers, named
parameters, and query chaining. For longer code, store a readable `.pq` or `.m`
file and use the M-code file input for evaluation and create/update. The source
filename does not need to match the Excel query name.

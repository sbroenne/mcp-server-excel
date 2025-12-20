# excel_powerquery - Server Quirks

**Action disambiguation**:

- create: Import NEW query using inline `mCode` (FAILS if query already exists - use update instead)
- update: Update EXISTING query M code + refresh data (use this if query exists)
- load-to: Loads to worksheet or data model or both (not just config change) - CHECKS for sheet conflicts
- unload: Removes data from ALL destinations (worksheet AND Data Model) - keeps query definition
- delete: Completely removes query AND all associated data (worksheet, Data Model connections)

**When to use create vs update**:

- Query doesn't exist? → Use create
- Query already exists? → Use update (create will error "already exists")
- Not sure? → Check with list action first, then use update if exists or create if new

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

- Using create on existing query → ERROR "Query 'X' already exists" (should use update)
- Using update on new query → ERROR "Query 'X' not found" (should use create)
- Calling LoadTo without checking if sheet exists (will error if sheet exists)
- Assuming unload only removes worksheet data → Also removes Data Model connections

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

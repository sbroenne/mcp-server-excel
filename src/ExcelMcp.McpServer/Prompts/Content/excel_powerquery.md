# excel_powerquery - Server Quirks

**Action disambiguation**:
- create: Import NEW query (FAILS if query already exists - use update instead)
- update: Update EXISTING query M code + refresh data (use this if query exists)
- load-to: Applies destination + refreshes (not just config change) - CHECKS for sheet conflicts
- unload: Removes data but keeps query definition (inverse of load-to)

**When to use create vs update**:
- Query doesn't exist? → Use create
- Query already exists? → Use update (create will error "already exists")
- Not sure? → Check with list action first, then use update if exists or create if new

**LoadTo with existing sheets (Issue #170 fix)**:
- If target worksheet already exists, LoadTo returns clear error
- User must delete existing sheet first using excel_worksheet action='Delete'
- Then retry LoadTo - ensures explicit user control over data deletion
- Example workflow: LoadTo fails → Delete sheet manually → Retry LoadTo succeeds

**Common mistakes**:
- Using create on existing query → ERROR "Query 'X' already exists" (should use update)
- Using update on new query → ERROR "Query 'X' not found" (should use create)
- Calling LoadTo without checking if sheet exists (will error if sheet exists)

**Server-specific quirks**:
- Validation = execution: M code only validated when data loads/refreshes
- connection-only queries: NOT validated until first execution
- refresh with loadDestination: Applies load config + refreshes (2-in-1)
- Single cell returns [[value]] not scalar

# excel_powerquery - Server Quirks

**Action disambiguation**:
- create vs import: create = atomic (import + load), import = M code only first
- load-to: Applies destination + refreshes atomically (not just config change)
- update-mcode vs update: update-mcode never refreshes, update may refresh
- update-and-refresh: Atomic code update + data refresh
- unload: Removes data but keeps query definition

**Server-specific quirks**:
- Validation = execution: M code only validated when data loads/refreshes
- connection-only queries: NOT validated until first execution
- refresh with loadDestination: Applies load config + refreshes (2-in-1)
- Single cell returns [[value]] not scalar

# excel_powerquery - Server Quirks

**Action disambiguation**:
- create: Import M code + load data in one operation (default: loads to worksheet)
- load-to: Applies destination + refreshes (not just config change)
- update: Updates M code + refreshes data (complete operation, keeps data fresh)
- unload: Removes data but keeps query definition (inverse of load-to)

**Server-specific quirks**:
- Validation = execution: M code only validated when data loads/refreshes
- connection-only queries: NOT validated until first execution
- refresh with loadDestination: Applies load config + refreshes (2-in-1)
- Single cell returns [[value]] not scalar

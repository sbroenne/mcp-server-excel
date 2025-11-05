# excel_querytable Tool

**Actions**: list, get, create-from-connection, create-from-query, refresh, refresh-all, update-properties, delete

**When to use excel_querytable**:
- Simple data imports from existing connections (no M code complexity)
- Load Power Query results to worksheets
- Legacy Excel workflows requiring QueryTable functionality
- Refresh data with synchronous pattern (guaranteed persistence)
- Use excel_connection for connection lifecycle management
- Use excel_powerquery for M code transformations
- Use excel_range for data access from QueryTables

**Server-specific behavior**:
- QueryTables use synchronous refresh (queryTable.Refresh(false)) for guaranteed persistence
- Single cell returns as 2D array in range data [[value]]
- RefreshImmediately defaults to true (immediate feedback after creation)
- QueryTables live on specific worksheets (not workbook-level)
- Each QueryTable has unique name within workbook

**Action disambiguation**:
- create-from-connection: Import data from existing connection (OLEDB, ODBC, Text, Web)
- create-from-query: Load Power Query results to worksheet (simpler than Power Query load-to)
- refresh: Synchronous single QueryTable refresh (guaranteed persistence)
- refresh-all: Synchronous refresh of all QueryTables in workbook
- update-properties: Modify refresh settings and formatting options

**Common mistakes**:
- Using async RefreshAll() instead of synchronous refresh → QueryTables won't persist
- Creating QueryTable without creating worksheet first → Must create sheet first
- Forgetting to specify connection/query name → Required for create actions
- Not refreshing after creation when refreshImmediately=false → Data won't appear

**Workflow optimization**:
- Creating multiple QueryTables? Use begin_excel_batch for better performance
- After creating QueryTable: Use excel_range to read data
- For Power Query: create-from-query is simpler than excel_powerquery load-to
- Common pattern: create-from-connection → refresh → read with excel_range

**Integration examples**:
- List connections: excel_connection list → create-from-connection
- List queries: excel_powerquery list → create-from-query
- Read data: create-from-query → excel_range get-values
- Update source: update-properties (change refresh settings) → refresh

# excel_connection Tool

**Related tools**:
- excel_powerquery - For creating new data connections (recommended)

**Actions**: list, view, create, test, refresh, delete, load-to, get-properties, set-properties

**When to use excel_connection**:
- Create and manage Excel connections (OLEDB, ODBC, TEXT, WEB)
- Refresh data from connection sources
- Delete connections you no longer need
- Load connection data to worksheets
- Configure connection properties (background refresh, auto-refresh, etc.)
- Use excel_powerquery for M query-based data connections

**Server-specific behavior**:
- TEXT and WEB connections: Can create programmatically using valid connection strings
- OLEDB/ODBC connections: Cannot be created via COM API (Excel limitation) - create in Excel UI first
- Connection types 3 and 4 (TEXT/WEB) may report inconsistently
- Delete removes connection and associated QueryTables
- Power Query connections automatically redirect to excel_powerquery tool

**Action disambiguation**:
- list: Show all connections in workbook
- view: Display connection details and properties
- create: Create new TEXT or WEB connection programmatically
- test: Verify connection without refreshing data
- refresh: Update data from connection source
- delete: Remove connection and its QueryTables (use excel_powerquery delete for Power Query connections)
- load-to: Load connection data to specified worksheet
- get-properties: Retrieve connection properties (background query, refresh settings, etc.)
- set-properties: Update connection properties (background query, refresh-on-open, save password, refresh period)

**Common mistakes**:
- Trying to delete Power Query connections → Use excel_powerquery delete instead
- Trying to create OLEDB/ODBC connections → Use Excel UI (Data → Get Data), then manage with this tool
- Not testing connection before refresh → Use test first to verify connectivity

**Workflow optimization**:
- For OLEDB/ODBC: Create in Excel UI → Manage/refresh/delete with this tool
- For TEXT/WEB: Can create programmatically with this tool
- Use test action before refresh to avoid errors
- Check connection properties with get-properties before modifying with set-properties

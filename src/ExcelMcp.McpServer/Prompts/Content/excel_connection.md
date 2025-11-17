# excel_connection Tool

**Related tools**:
- excel_powerquery - For creating new data connections (recommended)
- excel_querytable - QueryTables use connections for data loading

**Actions**: list, view, import, export, update-properties, test, refresh, delete, load-to, get-properties, set-properties

**When to use excel_connection**:
- Create and manage Excel connections (OLEDB, ODBC, TEXT, WEB)
- Refresh data from connection sources
- Import/export connection (.odc) files
- Use excel_powerquery for M query-based data connections

**Server-specific behavior**:
- All connection types: Can create programmatically using valid connection strings
- OLEDB/ODBC: Requires installed providers and valid connection strings
- TEXT connections: Requires valid file paths
- Connection types 3 and 4 (TEXT/WEB) may report inconsistently
- Passwords NOT exported by default (security)

**Action disambiguation**:
- list: Show all connections in workbook
- view: Display connection details
- test: Verify connection is working
- refresh: Update data from source
- import: Load connection from .odc file
- export: Save connection to .odc file

**Common mistakes**:
- Trying to create OLEDB connections → Use Excel UI or .odc files
- Expecting passwords in exports → Security prevents this
- Not testing connection before refresh → Use test first

**Workflow optimization**:
- Create connections in Excel UI → Manage with this tool
- Use test before refresh to avoid errors

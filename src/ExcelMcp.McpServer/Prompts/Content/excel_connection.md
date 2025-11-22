# excel_connection Tool

**Related tools**:
- excel_powerquery – For Power Query M sources, CSV/TEXT/WEB imports, or when an OLE DB provider is unavailable

**Actions**: list, view, create, test, refresh, delete, load-to, get-properties, set-properties

**When to use excel_connection**:
- Create and manage Excel connections (OLEDB, ODBC)
- Refresh data from connection sources
- Delete connections you no longer need
- Load connection data to worksheets
- Configure connection properties (background refresh, auto-refresh, etc.)
- Use excel_powerquery for Power Query/M-driven sources, CSV/TEXT/WEB imports, or when the target provider is missing

**Server-specific behavior**:
- OLEDB connections: Fully supported when the provider is installed (e.g., Microsoft.ACE.OLEDB.16.0, SQLOLEDB). Excel throws "Value does not fall within the expected range" if the provider is missing.
- ODBC connections: Supported; DSN or DSN-less connection strings work.
- TEXT/WEB connections: Creation routed to Power Query (use excel_powerquery create)
- DataFeed / Model: These types show up when workbook already has Power Query or Power Pivot connections. Manage/refresh them here; creation happens via excel_powerquery / Power Pivot UI.
- Connection types 3 and 4 (TEXT/WEB) may report inconsistently
- Delete removes connection and associated QueryTables
- Power Query connections automatically redirect to excel_powerquery tool
- Refresh/load-to actions time out after 5 minutes; if Excel is blocked (credentials, privacy dialog), you'll receive `SuggestedNextActions` instead of a hung session.

**Action disambiguation**:
- list: Show all connections in workbook
- view: Display connection details and properties
- create: Create OLEDB/ODBC connection (provider must exist). Use excel_powerquery for TEXT/WEB/Power Query scenarios.
- test: Verify connection without refreshing data
- refresh: Update data from connection source
- delete: Remove connection and its QueryTables (use excel_powerquery delete for Power Query connections)
- load-to: Load connection data to specified worksheet
- get-properties: Retrieve connection properties (background query, refresh settings, etc.)
- set-properties: Update connection properties (background query, refresh-on-open, save password, refresh period)

**Common mistakes**:
- Missing provider for OLEDB connection string → Install provider (e.g., ACE) or switch to ODBC/Power Query.
- Trying to create Power Query/TEXT/WEB connections → Use excel_powerquery instead.
- Not testing connection before refresh → Use test first to verify connectivity

**Workflow optimization**:
- Provide concrete provider connection strings (e.g., `Provider=Microsoft.ACE.OLEDB.16.0;...`).
- Use test before refresh to surface provider/network errors earlier.
- Route CSV/TEXT/WEB scenarios to excel_powerquery (faster automation, no provider requirement).
- Check properties with get-properties before modifying via set-properties.

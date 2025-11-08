# excel_powerquery Tool

**⚠️ BEFORE CALLING**: Check powerquery_import.md elicitation for complete parameter checklist

**Related tools**:
- excel_batch - Use for 2+ operations (75-90% faster)
- excel_datamodel - For DAX measures on imported data
- excel_connection - For managing underlying data connections
- excel_table - For data already in Excel worksheets

**Actions**: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config, errors, test, load-to, sources, peek, eval, **create, update-mcode, unload, update-and-refresh, refresh-all**

**✨ NEW: Atomic Operations (Recommended)**:
- **create**: Import M code + load to a worksheet in ONE operation
  - Example: `create(queryName, mCodeFile, loadMode='worksheet')` → Query ready to use!
- **load-to**: Convert connection-only query to loaded state (set destination + refresh)
  - Example: `load-to(queryName, loadDestination='worksheet')` → Connection-only → worksheet!
- **update-and-refresh**: Update M code + refresh data atomically
  - Example: `update-and-refresh(queryName, mCodeFile)` → Code updated + data current!
- **refresh-all**: Batch refresh all queries in workbook
  - Example: `refresh-all(excelPath)` → All queries refreshed in one call!
- **update-mcode**: Stage M code changes without refreshing (faster for iterative development)
  - Example: `update-mcode(queryName, mCodeFile)` → Code updated, data unchanged
- **unload**: Convert loaded query to connection-only (inverse of load-to)
  - Example: `unload(queryName)` → Query definition kept, data removed

**When to use atomic operations**:
- **create** → New query with data
- **load-to** → Convert connection-only to loaded: Apply destination + refresh in one call
- **update-and-refresh** → Production updates: M code + data current in one atomic operation
- **refresh-all** → Batch data refresh: All queries updated together
- **update-mcode** → Iterative development: Update code without waiting for refresh
- **unload** → Remove data but keep query: Useful for cleanup or switching to connection-only

**When to use excel_powerquery**:
- External data sources (databases, web APIs, files, SharePoint)
- Power Query M code transformations
- Data refresh workflows
- Use excel_table for data already in Excel worksheets
- Use excel_datamodel for DAX measures after loading to data model

**Server-specific behavior**:
- import DEFAULT: loadDestination='worksheet' (validates M code by executing it)
- create DEFAULT: loadMode='worksheet' (atomic import + load)
- loadDestination='data-model': Loads to Power Pivot (ready for DAX, NOT visible in worksheet)
- loadDestination='both': Visible in worksheet AND available for DAX
- loadDestination='connection-only': M code imported but NOT executed/validated
- refresh with loadDestination: Applies load config + refreshes in ONE call (avoids set-load + refresh)

**Load destination guide**:
- 'worksheet': Users see data in Excel, no DAX capability
- 'data-model': Ready for DAX measures, users don't see data
- 'both': Best of both worlds (visibility + DAX)
- 'connection-only': Advanced only, no validation

**Action disambiguation**:
- **create**: Add new query + load data atomically (RECOMMENDED for new queries)
- import: Add new query from .pq file (use loadDestination parameter)
- **load-to**: Convert connection-only query to loaded state (RECOMMENDED over set-load + refresh)
- **update-and-refresh**: Update M code + refresh atomically (RECOMMENDED for production updates)
- **update-mcode**: Update M code only (no refresh, faster for iterative dev)
- update: Modify existing query M code (preserves load config, legacy)
- **refresh-all**: Refresh all queries in workbook (RECOMMENDED for batch refresh)
- refresh: Refresh data from source (optionally apply loadDestination, legacy)
- **unload**: Remove data, keep query definition (inverse of load-to)
- set-load-to-table: Change existing query to load to worksheet only (legacy - use load-to)
- set-load-to-data-model: Change existing query to load to Power Pivot (legacy - use load-to)
- set-load-to-both: Change existing query to load to both destinations (legacy - use load-to)
- set-connection-only: Prevent data loading (M code only)

**Common mistakes**:
- Forgetting loadDestination on import → defaults to 'worksheet', not data model
- Using set-load-to-table then trying excel_table add-to-datamodel → Use loadDestination='data-model' or 'both' instead
- Two-step workflow (import + set-load) → Use **create** action instead (atomic)
- Two-step workflow (set-load + refresh) → Use **load-to** action instead (atomic)
- Two-step workflow (update + refresh) → Use **update-and-refresh** action instead (atomic)
- Refreshing queries one-by-one → Use **refresh-all** for batch operations
- Expecting connection-only queries to validate → M code only validated when executed

**Workflow optimization**:
- Multiple imports? Use begin_excel_batch first (75-90% faster)
- New query with data? Use **create** action (replaces import + load-to)
- Connection-only query needs data? Use **load-to** action (replaces set-load + refresh)
- Production M code updates? Use **update-and-refresh** (replaces update + refresh)
- Multiple queries to refresh? Use **refresh-all** (single call, faster)
- After loading to data model: Use excel_datamodel for DAX measures/relationships
- Changing load destination: No need to delete/recreate, just use set-load-* or **load-to** actions

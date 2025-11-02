# excel_powerquery Tool

**Actions**: list, view, import, export, update, refresh, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config, errors, test, load-to, sources, peek, eval

**When to use excel_powerquery**:
- External data sources (databases, web APIs, files, SharePoint)
- Power Query M code transformations
- Data refresh workflows
- Use excel_table for data already in Excel worksheets
- Use excel_datamodel for DAX measures after loading to data model

**Server-specific behavior**:
- import DEFAULT: loadDestination='worksheet' (validates M code by executing it)
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
- import: Add new query from .pq file (use loadDestination parameter)
- update: Modify existing query M code (preserves load config)
- refresh: Refresh data from source (optionally apply loadDestination)
- set-load-to-table: Change existing query to load to worksheet only
- set-load-to-data-model: Change existing query to load to Power Pivot
- set-load-to-both: Change existing query to load to both destinations
- set-connection-only: Prevent data loading (M code only)

**Common mistakes**:
- Forgetting loadDestination on import → defaults to 'worksheet', not data model
- Using set-load-to-table then trying excel_table add-to-datamodel → Use loadDestination='data-model' or 'both' instead
- Two-step workflow (set-load + refresh) → Use refresh with loadDestination parameter (one call)
- Expecting connection-only queries to validate → M code only validated when executed

**Workflow optimization**:
- Multiple imports? Use begin_excel_batch first (75-90% faster)
- After loading to data model: Use excel_datamodel for DAX measures/relationships
- Changing load destination: No need to delete/recreate, just use set-load-* actions

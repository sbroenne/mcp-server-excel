---
"excelmcp": patch
---

Power Query `refresh-all` now works on workbooks that contain parameter or connection-only queries: it skips them, refreshes every loaded query, and keeps going when one query fails. The result lists which queries were refreshed, skipped, and failed (with each failure's error), and reports `success: false` when any query failed. `excelcli` now exits with code 1 whenever a command's result reports `success: false`, matching how the MCP Server flags the same result as an error.

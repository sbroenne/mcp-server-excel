---
"excelmcp": patch
---

Fix Power Query and workbook `refresh-all` reporting PivotTables as refreshed while their totals are still out of date. Each report now refreshes individually, including when several reports use separate caches for the same source. Applies to both the CLI and MCP Server.

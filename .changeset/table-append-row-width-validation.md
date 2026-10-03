---
"excelmcp": patch
---

Validate every appended table row against the table's column count before writing cells. MCP and excelcli now reject incomplete or oversized rows without silently dropping values, expanding the table, or changing calculation mode.

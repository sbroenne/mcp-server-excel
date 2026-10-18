---
"excelmcp": patch
---

External OLAP member searches no longer report a zero or undersized total when the provider returns more members. Inconsistent totals are left unknown while member results and paging remain available in both MCP and the CLI.

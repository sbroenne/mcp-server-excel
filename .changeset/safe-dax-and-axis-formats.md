---
"excelmcp": patch
---

Remove regional separator rewriting when creating and updating DAX measures, fixing expressions such as DIVIDE(SUM(SalesTable[Amount]), 1000). Native DAX separators are preserved. On decimal-comma Windows, spaces are added around commas that touch numbers so Excel keeps the formula's meaning, and the result message reports the adjustment. Remote formatting remains opt-in. Fix General and named-date formats for cells, and preserve chart-axis currency symbols and date/time meaning when applying US format codes through MCP or excelcli.

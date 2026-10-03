---
"excelmcp": patch
---

Remove regional separator rewriting when creating and updating DAX measures, fixing expressions such as DIVIDE(SUM(SalesTable[Amount]), 1000). Formulas are passed to Excel unchanged unless remote formatting is explicitly requested. Fix General and named-date formats for cells, and preserve chart-axis currency symbols and date/time meaning when applying US format codes through MCP or excelcli.

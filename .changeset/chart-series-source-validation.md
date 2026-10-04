---
"excelmcp": patch
---

Resolve chart series source ranges before adding a series. Missing sheets and invalid ranges now fail without adding a partial series or treating the reference as a literal chart value through MCP or excelcli.

Resolve unqualified series sources from the chart worksheet and sheet-qualified sources from the chart workbook, independent of the active worksheet or workbook.

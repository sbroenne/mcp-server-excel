---
"excelmcp": patch
---

Resolve chart series source ranges before adding a series. Missing sheets and invalid ranges now fail without adding a partial series or treating the reference as a literal chart value through MCP or excelcli.

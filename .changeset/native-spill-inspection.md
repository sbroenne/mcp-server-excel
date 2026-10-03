---
"excelmcp": minor
---

Add `range` `get-spill-info` to inspect every requested cell's native
dynamic-array source, result, ordinary, or blocked state through both the MCP
Server and `excelcli`. Include actual source formulas and established result
extents without recalculating, changing selection, sampling, or inventing
blocked extents. Unsupported Excel sessions fail explicitly.

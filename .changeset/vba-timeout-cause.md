---
"Sbroenne.ExcelMcp.McpServer": patch
"Sbroenne.ExcelMcp.CLI": patch
---

`vba run` now reports its requested execution timeout as a timeout rather than a
cancellation. Anonymous operation analytics also distinguish timeouts from
cancellations.

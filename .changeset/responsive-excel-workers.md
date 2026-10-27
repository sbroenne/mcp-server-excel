---
"excelmcp": patch
---

Move Excel work out of the CLI background service and MCP controller so status and tray controls do not wait behind workbook operations. Bound service stopping and connection cleanup without saving, discarding, or ending Excel work to escape an expired wait. Keep sequential batch commands and JSON results, with live progress written separately.

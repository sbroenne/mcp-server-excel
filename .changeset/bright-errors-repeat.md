---
"excelmcp": major
---

**Consistent error output**: MCP and `excelcli` failure replies now use only `errorMessage` for the error text. Update consumers that read the removed `error` field.

---
"excelmcp": major
---

MCP tools that target an open workbook now use `workbook_session_id` instead of `session_id` in inputs and results. This works around a Claude Desktop bridge issue that can drop inputs named `session_id`. CLI session names are unchanged.

---
"excelmcp": patch
---

Workbook close responses now automatically include a compact `closeTimingsMs` breakdown, including work done in background processes. Both the CLI and MCP return the measurements directly to the calling agent, with no environment setting or local timing files needed. Saving, closing, required cleanup, and time limits are unchanged.

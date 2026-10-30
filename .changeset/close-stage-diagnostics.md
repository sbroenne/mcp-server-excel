---
"excelmcp": patch
---

CLI and MCP workbook close responses automatically return a compact `closeTimingsMs` breakdown to the calling agent, without environment settings or timing files.

Closing no longer forces general memory cleanup or waits for the Excel process to disappear after any requested save, confirmed workbook close, any appropriate Quit, and release of the Excel objects we used. An empty background copy of ExcelMcp finishes through the same internal reply without extra status or stop requests, while saving rules, protection of other workbooks, and recovery after failures remain unchanged.

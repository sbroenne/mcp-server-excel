---
"excelmcp": patch
---

When Excel cannot provide the connection needed for OLAP discovery or PivotTable selection reads, ExcelMcp now checks for blocking Excel dialogs automatically and includes readable dialog text or an explicit check result in the error. Recovery suggests completing any needed sign-in in desktop Excel and retrying before considering a refresh, which still requires separate permission and may change workbook data.

Errors caused by a confirmed Excel dialog now include its readable text immediately for Save, Save As, closing, shutdown, refresh controls, and connection account settings; Save As also checks for blocking dialogs before waiting on an attached workbook, without answering dialogs or saving or discarding edits.

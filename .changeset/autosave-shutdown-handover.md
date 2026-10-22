---
"excelmcp": minor
---

Opening workbooks now preserves Excel AutoSave and reports its status. Use the new workbook `set-auto-save` action in MCP or excelcli to change it explicitly; enabling it can save pending edits.

Stopping ExcelMcp no longer saves pending edits automatically. Saved workbooks it opened close, while unsaved ones stay open and visible for you to review. Attached workbooks remain untouched. Save and close explicitly when a completed job should leave no workbook open.

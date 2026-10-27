---
"excelmcp": minor
---

Add four actions in ExcelMcp and excelcli to check workbook opening restrictions, allow editing, request access to data connections, or enable VBA macros (code stored in the workbook). Each change needs separate approval; allowing editing or connections also needs approval for Excel to remember that the file is trusted.

Existing workbooks now open with VBA macros disabled, including .xlsm files. Enabling them requires approval to close and reopen a saved workbook on this computer, which can run code as it opens; unsaved changes are never saved or discarded automatically.

---
"excelmcp": patch
---

Fix saving and closing workbooks after Data Model queries. ExcelMcp now avoids a screen-update change that could make Excel mark a saved workbook as changed again. Save failures still leave the session open and report the error.

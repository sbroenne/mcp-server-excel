---
"excelmcp": patch
---

Fix Power BI/MSOLAP connection inspection and refresh when Excel does not expose background-query settings. Save, Save As, and close now check Excel's live refresh state, keep unfinished work open, and report busy errors without incorrectly blaming a file lock.

Require confirmed Excel process ownership during startup, retrying temporary identity-capture failures and rejecting startup before opening or creating a workbook if ownership cannot be confirmed.

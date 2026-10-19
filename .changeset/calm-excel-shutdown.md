---
"excelmcp": patch
---

Improve Excel cleanup after reading Data Model tables or creating or updating measures, including failed changes. Read-only workbook checks no longer try to save during shutdown, and Excel has more time to finish closing naturally.

`excelcli service stop` now waits for the shutdown reply until cancelled instead of giving up after one minute. Excel's own save and shutdown time limits still apply.

---
"excelmcp": patch
---

When an Excel operation times out with a confirmed blocking dialog, report that Excel was waiting for input without reading or exposing dialog titles or contents. Timeout recovery retains the session and unsaved workbook without saving, closing, or rolling back the operation.

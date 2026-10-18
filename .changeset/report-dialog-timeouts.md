---
"excelmcp": patch
---

When an Excel operation times out with a confirmed blocking dialog, report that Excel was waiting for input and include its readable title and message labels automatically for troubleshooting. Unavailable or partial text is marked; input values are not read and the dialog is not answered. Timeout recovery retains the session and unsaved workbook without saving, closing, or rolling back the operation.

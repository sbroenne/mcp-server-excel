---
"excelmcp": patch
---

Move Excel work out of the CLI background service and MCP controller so status and tray controls do not wait behind workbook operations. Bound service stopping and connection cleanup without saving, discarding, or ending Excel work to escape an expired wait. Keep sequential batch commands and JSON results, with live progress written separately.

Keep failed or unfinished cross-file worksheet transfers in a reported recovery session owning both workbooks. Preserve their exact partial state and original error without repeating the transfer or discarding edits. An explicit recovery save writes remaining changes in both workbooks without saving unchanged files again.

---
"excelmcp": patch
---

Validate worksheet names against Microsoft's documented restrictions before mutation. If Excel rejects a name beyond those checks after creating or copying a sheet, return an error identifying the sheet left in the workbook rather than removing it.

Reject blank existing worksheet names as invalid input before dispatching rename, copy, copy-to-file, or move-to-file operations.

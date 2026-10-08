---
"excelmcp": patch
---

Retain the workbook and service when Excel rejects the final shutdown save as busy, preserving the original COM error for inspection and retry. Remove sessions whose Excel process has exited from the daemon's idle count so abandoned sessions do not prevent idle shutdown.

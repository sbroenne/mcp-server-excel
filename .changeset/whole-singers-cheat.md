---
"excelmcp": patch
---

Improve Excel session safety and resource cleanup, prevent queued work from running after it times out, and handle single-cell ranges correctly. Connection inspection now hides stored credentials, cancelled refreshes no longer report success, and workbook-open errors retain their actual cause.

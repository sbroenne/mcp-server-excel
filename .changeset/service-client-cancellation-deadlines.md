---
"excelmcp": patch
---

Keep caller cancellation distinct from service timeouts and connection failures, including when the pipe disconnects at the same time. Preserve unlimited connection waits while respecting any finite total request deadline.

If opening or creating a workbook finishes just as the caller cancels, reveal the completed workbook before reporting cancellation. Do not save, close, or undo its changes.

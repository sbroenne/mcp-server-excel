---
"excelmcp": minor
---

**Incremental CLI batches:** `batch --stream` processes one JSON command at a time from stdin and returns each result immediately, so callers can inspect a result before sending the next command through the same client. Existing file-based batch input remains compatible. A failed close now retains the active session selection for inspection.

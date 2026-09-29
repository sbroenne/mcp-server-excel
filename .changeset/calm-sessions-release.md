---
"excelmcp": patch
---

**Safer Excel sessions and connection output**: Queued operations now expire
without running later or closing a healthy session, connection details redact
credentials, one-cell reads return correctly, refresh cancellation is reported,
and Excel COM objects are released more reliably.

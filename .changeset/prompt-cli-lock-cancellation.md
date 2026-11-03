---
"excelmcp": patch
---

Cancelling CLI startup or `service stop` now promptly stops waiting for another startup or stop request. Existing wait limits and protection against duplicate services are unchanged.

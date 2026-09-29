---
"excelmcp": patch
---

Retry transient Windows sharing failures while replacing CLI daemon tracking
state so concurrent cleanup and startup do not leave a locked runtime behind.

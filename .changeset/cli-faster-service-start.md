---
"excelmcp": patch
---

The CLI starts its background service about 4 seconds faster, and `service stop` returns about 1.5 seconds faster when the service is not running. Neither waits for a connection timeout when the service is not there.

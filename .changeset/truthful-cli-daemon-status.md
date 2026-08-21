---
"excelmcp": patch
---

**Truthful CLI daemon status** (#785): `service status` now distinguishes
stopped, running, and unresponsive daemons, while `session list` returns an empty
list only for confirmed empty states. Both commands probe the configured service
before consulting daemon mutex state, so externally hosted services remain
visible. Shared control-command and startup readiness timeouts tolerate slow
daemon startup without converting transport failures into successful stopped or
empty results.

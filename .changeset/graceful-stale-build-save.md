---
"excelmcp": patch
---

**Stale-build cleanup now preserves unsaved CLI session changes when a shutdown reply is lost.** The cleanup client gives an already-shutting-down daemon and its tracked Excel processes time to save and exit before using pipe-scoped forced cleanup.

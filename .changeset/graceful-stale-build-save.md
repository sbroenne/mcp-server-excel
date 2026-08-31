---
"excelmcp": patch
---

**Stale-build cleanup now preserves unsaved CLI session changes when a shutdown reply is lost.** Session and pipe cleanup use one exact-process exit policy, giving an already-shutting-down daemon and its tracked Excel processes time to save and exit before bounded forced cleanup.

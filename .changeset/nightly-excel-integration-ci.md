---
"excelmcp": patch
---

**Automated Excel integration testing.** A cost-optimized self-hosted Windows runner now executes the real Excel integration suite nightly, including VBA and session tests. The runner starts on demand and is deallocated after testing, closing the previous gap where Excel-dependent behavior was validated only on a maintainer workstation.

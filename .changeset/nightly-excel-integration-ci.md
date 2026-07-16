---
"excelmcp": patch
---

**Automated Excel integration testing.** A cost-optimized self-hosted Windows runner now executes the real Excel integration suite for ready pull requests, including VBA and session tests. Targeted manual runs support surgical feature validation during development, while only the full-suite check satisfies the merge gate. The runner starts on demand and is deallocated after testing, and its setup downloads are verified before execution. Formula reads now report correct worksheet coordinates and actionable suggestions for cell errors, and locked or invalid workbook paths are rejected before Excel starts.

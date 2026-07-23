---
"excelmcp": patch
---

**Release automation: make the changelog commit-back reliable.** The post-release step that writes `CHANGELOG.md` back to `main` now opens a short-lived PR and merges it with an admin bypass, instead of pushing directly to `main`, fixing a case where the direct push was unexpectedly rejected.

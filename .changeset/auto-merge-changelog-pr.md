---
"excelmcp": patch
---

**Release automation: auto-merge the changelog PR.** The post-release step that opens the `chore/changelog-vX` PR now also merges it (queued auto-merge, falling back to an immediate squash merge). Previously the PR was only created and left open until a maintainer merged it by hand, which caused several releases to sit with a stale/missing CHANGELOG on `main`.

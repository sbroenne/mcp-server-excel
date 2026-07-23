---
"excelmcp": patch
---

**Release automation: reliably commit the changelog back to `main`.** The post-release step now pushes the compiled `CHANGELOG.md` update directly to `main` using an admin `RELEASE_PAT`, instead of opening a `chore/changelog-vX` PR. On this user-owned repo the GitHub Actions bot can't be a branch-protection bypass actor, so that PR could never satisfy the required status checks and piled up open — leaving several releases with a stale/missing CHANGELOG on `main`. The direct push (as a ruleset bypass actor) removes the stuck-PR failure mode entirely.

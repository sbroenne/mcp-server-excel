---
"excelmcp": patch
---

**Release automation: make the changelog commit-back reliable.** The post-release step that writes `CHANGELOG.md` back to `main` now opens a short-lived PR and merges it with an admin bypass, instead of pushing directly to `main`. A real release proved that GitHub can reject a raw `git push` to a ruleset-protected branch from inside a GitHub Actions runner even when the same admin PAT successfully bypasses the identical ruleset when used interactively — the admin-merge API path doesn't have that restriction.

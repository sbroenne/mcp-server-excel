---
"excelmcp": patch
---

**Faster commits for docs-only changes.** The pre-commit hook now treats the `gh-pages/` documentation website as docs and skips the Release build, smoke tests and all release-packaging gates when a commit touches only documentation (Markdown, `docs/`, `gh-pages/`, changesets). Code commits still run the full validation suite, so nothing that ships is left unchecked — documentation edits just no longer wait minutes for binary/packaging gates that cannot be affected by them.

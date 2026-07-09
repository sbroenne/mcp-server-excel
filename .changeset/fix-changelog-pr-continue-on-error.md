---
"excelmcp": patch
---

**Release automation hardening**: The post-release step that opens a PR to commit the compiled `CHANGELOG.md` no longer silently swallows failures. During the first live run of the new changesets-based release pipeline, this step failed (the repo didn't allow Actions to create pull requests) but was marked as a passing step, which is exactly the kind of silent failure the new pipeline was built to eliminate. The repo setting has been fixed and the step now fails the release run loudly if it can't create the PR.

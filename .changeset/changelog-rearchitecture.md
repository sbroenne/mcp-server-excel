---
"excelmcp": patch
---

**Changelog generation now uses changesets** (#698): Each PR adds a small, human-written note describing what changed for users, and these notes are compiled automatically into `CHANGELOG.md` and the GitHub Release notes when a new version ships. This replaces the old manual process, which had let several releases' worth of changes sit mislabeled as "Unreleased" for months. The changelog itself has also been cleaned up — the mislabeled entries were consolidated and condensed into clearer, less technical summaries.

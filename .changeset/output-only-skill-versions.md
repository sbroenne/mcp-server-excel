---
"excelmcp": patch
---

**Accurate skill versions in every distribution** (#791): plugin, Agent Skills ZIP,
and VS Code extension builds now stamp their resolved package version into generated
skill metadata instead of copying a stale source `VERSION` file. Manual distributable
builds must pass `-Version`, preventing silently mislabeled packages.

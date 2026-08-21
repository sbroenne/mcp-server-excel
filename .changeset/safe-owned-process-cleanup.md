---
"excelmcp": patch
---

**Safer CLI cleanup** (#789): builds and test workflows now stop only the CLI daemon and Excel processes owned by their selected pipe, preserving unrelated Excel sessions. Daemon lifecycle locks use disjoint, case-insensitive hashed pipe identities so case variants, suffixes, special characters, and long names cannot collide. When cleanup sources are newer than the installed build output, pre-build cleanup uses an isolated current client so the owned daemon cannot lock the rebuild.

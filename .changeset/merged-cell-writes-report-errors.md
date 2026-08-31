---
"excelmcp": patch
---

**Merged-cell writes now fail clearly instead of reporting false success** (#831). `set-values` and `set-formulas` identify affected merged ranges and explain whether to write to the top-left cell or unmerge the range.

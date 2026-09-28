---
"excelmcp": patch
---

**Unstyled tables can be listed and read again**: The `table list` and `table read`
actions now handle Excel tables that have no table style applied.
No-style values are returned as an empty string, while applied style names are preserved.

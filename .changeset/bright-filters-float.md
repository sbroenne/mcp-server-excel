---
"excelmcp": patch
---

PivotTable filters now work on OLAP/Data Model row and column fields, including caption, value, and top/bottom filters. Date filters are also supported when Excel and the provider recognize the source date level. Value filters and sorting accept either the displayed measure caption or its `[Measures]` name. Filter commands accept placed hierarchy and level names, including the level names returned by filter reads. Semantic snapshots report member-property filter criteria as unavailable instead of claiming complete filter state.

---
"excelmcp": patch
---

PivotTable filters now work on Data Model row and column fields, including caption, value, and top/bottom filters. Value filters and sorting accept either the displayed measure caption or its `[Measures]` name. OLAP date-filter requests return a clear unsupported-operation error. Filter commands accept placed hierarchy and level names, including the level names returned by filter reads. Semantic snapshots report member-property filter criteria as unavailable instead of claiming complete filter state.

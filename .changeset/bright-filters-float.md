---
"excelmcp": patch
---

PivotTable filters now work on Data Model row and column fields, including caption, value, and top/bottom filters. Value filters and sorting accept either the displayed measure caption or its `[Measures]` name. Excel does not support native date filters on OLAP PivotTables, so those requests now return a clear error.

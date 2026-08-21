---
"excelmcp": patch
---

**Exact Power Query and connection cleanup** (#786, #796, #797): Power Query
load detection, refresh, unload, delete, Data Model lookup, and evaluate cleanup
now use exact case-insensitive mashup `Location` identity instead of substring or
display-name matching. Connection delete/load-to removes only QueryTables owned
by the exact WorkbookConnection, preserving similarly named and unrelated
workbook objects. Queries loaded to both a worksheet and the Data Model now
refresh both destinations instead of leaving model data stale. Evaluate removes
Excel-generated `Connection`/`Connection1` artifacts before save and returns an
actionable error if cleanup fails.

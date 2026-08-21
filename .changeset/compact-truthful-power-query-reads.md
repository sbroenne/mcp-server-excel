---
"excelmcp": major
---

**Compact, truthful Power Query reads** (#787, #800): `powerquery list` now
returns bounded M previews and exact worksheet/Data Model load state without
serializing full formulas. Use `powerquery view` for full M code. List inspection
errors now fail explicitly instead of silently omitting queries. The public
`PowerQueryInfo.Formula` getter and setter remain available for source and binary
compatibility, but are obsolete and excluded from list JSON.

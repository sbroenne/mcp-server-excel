---
"excelmcp": minor
---

**Safer table creation** (#838): Preview merged cells, header problems, nearby excluded columns, formula sorting risks, and the effective table range before creating a table. Table creation now blocks deterministic problems while leaving uncertain warnings for review. Large-range merge discovery is bounded, and formula risk analysis reports when an oversized range was skipped.

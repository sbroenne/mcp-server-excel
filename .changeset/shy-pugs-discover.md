---
"excelmcp": minor
---

External OLAP connections can now list dimensions, hierarchies, levels, and bounded member searches using Excel's existing connection session.

Member searches page in provider unique-name order over a bounded list of up to 10,000 entries. Larger levels fail explicitly instead of scanning without a limit. Paging tokens keep fixed-size fingerprints of the search settings and member list, so long search text does not produce an unusable token and a changed list cannot silently skip or repeat entries.

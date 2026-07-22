---
"excelmcp": minor
---

**Read existing conditional formatting rules** (#730): the `conditionalformat` tool now supports `list-rules` (per range) and `list-worksheet-rules` (entire sheet). Both return each rule's type, operator, formulas, applies-to range, priority, and formatting (interior/font/borders) with colors as `#RRGGBB` hex strings, in priority order — enabling round-trip safety, debugging, migration, and audit workflows before modifying or clearing rules.

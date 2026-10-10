---
"excelmcp": minor
---

External OLAP connections can now list dimensions, hierarchies, levels, and bounded member searches using Excel's existing connection session.

`pivottable_field` `add-value-field` now adds measures from an external OLAP cube (for example SQL Server Analysis Services) to a PivotTable. Previously it only found measures in the workbook's own Data Model and failed for server cubes.

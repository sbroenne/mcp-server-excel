---
"excelmcp": patch
---

Fix adding existing external cube measures to PivotTables, including measure names containing dots or escaped closing brackets. Explain which field changes must be made on the cube server rather than in the workbook Data Model.

Return actual OLAP field captions and positions, and apply automatic subtotals using hierarchy names. Preserve compound DAX formulas when changing an aggregation instead of silently replacing their calculation.

Report completed and uncertain hierarchy levels when a subtotal update fails, without claiming that partial changes were rolled back.

---
"excelmcp": patch
---

**Localized Excel table names**: Table operations now accept non-ASCII names
that Excel creates or allows, such as `表1` and `テーブル1`, instead of
rejecting them before the workbook is checked.

Rejected names are checked in a temporary Excel workbook before creating a table,
preserving source formulas, formatting, and headerless data. Preflight reports
the same rejected names as blockers.

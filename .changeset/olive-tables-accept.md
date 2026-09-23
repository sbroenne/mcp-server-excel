---
"excelmcp": patch
---

**Localized Excel table names**: Table operations now accept non-ASCII names
that Excel creates or allows, such as `表1` and `テーブル1`, instead of
rejecting them before the workbook is checked.

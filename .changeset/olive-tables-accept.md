---
"excelmcp": patch
---

**Localized Excel table names**: Table operations now accept non-ASCII names
that Excel allows, such as `表1` and `テーブル1`, instead of rejecting them
before Excel checks the name.

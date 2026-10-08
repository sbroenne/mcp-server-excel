---
"excelmcp": patch
---

**Formula blocks can mix in numbers, true/false, and blanks** ([#1072](https://github.com/sbroenne/mcp-server-excel/issues/1072)): `range set-formulas` and `range validate-formulas` now accept JSON numbers, `true`/`false`, and `null` (an empty cell) next to formulas and text, both inline in `formulas` and in a `formulasFile`. Numbers keep their exact value in any Excel language, and `validate-formulas` treats these constants as valid instead of reporting them as errors.

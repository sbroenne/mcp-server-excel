---
"excelmcp": patch
---

**`set-values` no longer blanks other cells when one value starts with `=`** ([#1065](https://github.com/sbroenne/mcp-server-excel/issues/1065)). Writing a mix of numbers, text, dates, and formulas now keeps every value: only the cells starting with `=` become formulas. Before, every other cell in the range was silently cleared while the command reported success.

**`get-formulas` no longer reports text that starts with `=` as a formula.** Cells holding text such as `'=abc` now return an empty formula and their text as the value. The `set-values` result message now reads `Wrote N formula(s); other cells kept as values` when formulas are written.

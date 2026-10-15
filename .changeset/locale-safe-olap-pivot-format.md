---
"excelmcp": patch
---

`pivottablefield set-field-format` now applies numeric and custom currency
formats correctly on non-US locales without changing the shared Data Model
measure. It returns the selected value field's displayed text, including in
PivotTables with multiple measures. OLAP measures can also be selected by their
exact caption.

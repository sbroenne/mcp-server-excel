---
"excelmcp": patch
---

Charts can now use data made of separate cell blocks on one sheet, such as labels in M4:M29 and values in O4:S29 (`M4:M29,O4:S29`), when creating a chart or changing its source. Bad input for charts, drawing objects, PivotTables, slicers, timelines and conditional formats is now rejected before anything is added to the workbook. If Excel still fails after creating the object, the error now names what was left behind and where, instead of showing a bare Excel error code.

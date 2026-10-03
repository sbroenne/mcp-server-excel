---
"excelmcp": patch
---

**DAX measures on decimal-comma computers** (#978): Creating or updating a Data Model measure no longer fails or silently changes numbers such as `-1, MONTH` into `-1.` when Windows uses a comma as the decimal mark. Spaces are added around commas that touch numbers, which keeps the DAX meaning, and the result message says so. Measure readback now returns the DAX stored in the model, with decimal points, after a workbook is reopened.

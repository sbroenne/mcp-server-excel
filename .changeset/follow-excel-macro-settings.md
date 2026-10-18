---
"excelmcp": minor
---

ExcelMcp now follows your own Excel macro settings instead of lowering macro security. A saved `.xlsm` outside a trusted location opens with its macros blocked: `Workbook_Open` no longer runs automatically, and `vba run` fails with Excel's own "macros disabled" message plus guidance on how to allow macros for the file. You can still list, view, and edit the code. To run macros in such a file, store it in a trusted location or sign its VBA project, then open it again.

---
"excelmcp": minor
---

Treat Excel workbooks as opaque files across production, tests, and scripts.
macOS `.xlsx` creation now copies an intact Excel-authored template, while
`.xlsm` creation is explicitly unsupported until an authentic macro-enabled
template is available. File validation no longer parses workbook package
internals or claims structural validity before Excel opens the file.

Report every macOS Power Query action as an explicit API limitation. Apple
Events and Office.js do not expose the required lifecycle contract, and
ExcelMcp neither inspects workbook packages nor ships a VBA add-in as a
fallback. Add a pre-commit audit that rejects direct workbook ZIP, OOXML,
relationship, and DataMashup access outside non-workbook distribution
packaging.

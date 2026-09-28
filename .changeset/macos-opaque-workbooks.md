---
"excelmcp": minor
---

Treat Excel workbooks as opaque files across production, tests, and scripts.
macOS `.xlsx` creation now copies an intact Excel-authored template, while
`.xlsm` creation is explicitly unsupported until an authentic macro-enabled
template is available. File validation no longer parses workbook package
internals or claims structural validity before Excel opens the file.

Route every macOS Power Query action exclusively through the optional trusted
helper. Helper 1.4.0 returns complete, validated list, view, and load metadata;
all actions remain blocked by default until exact public CLI and MCP acceptance
proves them. Add a pre-commit audit that rejects direct workbook ZIP, OOXML,
relationship, and DataMashup access outside non-workbook distribution
packaging.

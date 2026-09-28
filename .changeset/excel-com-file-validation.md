---
"excelmcp": patch
---

**Excel-backed file validation**: `file test` and `excelcli session test` now
validate ordinary workbooks by briefly opening them read-only in Excel instead
of inspecting internal workbook XML. The validation open honors the existing
MCP timeout, and the CLI command now supports the same `--timeout` option.

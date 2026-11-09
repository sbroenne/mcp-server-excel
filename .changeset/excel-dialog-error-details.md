---
"excelmcp": patch
---

Include a confirmed blocking Excel dialog's title and readable message in busy errors and commands refused after an earlier wait expired, in both excelcli and MCP. Preserve available dialog titles when the dialog is not responding, without changing timeouts, answering prompts, or discarding workbook edits.

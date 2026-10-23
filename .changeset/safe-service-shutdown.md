---
"excelmcp": patch
---

Stopping the service no longer fails or gets stuck. Exiting the CLI tray after you closed a workbook in Excel no longer shows an error. `excelcli service stop`, the tray's Exit, and MCP Server shutdown now always stop the service. Saved workbooks are closed, and unsaved ones stay open and visible without saving. If a workbook is busy, waiting for a prompt, or does not answer within the shutdown wait (at most 30 seconds), that Excel copy is left exactly as it is, made visible, and reported by process number: `service stop` returns it as `shutdownWarning`, the tray shows it, and the MCP Server writes it to its log. Nothing is saved, discarded, or force-stopped. Saving a workbook that was closed in Excel now fails with a clear message instead of an unclear Excel error.

---
"excelmcp": patch
---

Report unexpected Excel errors during optional property reads instead of silently returning missing data. Failed results now remain failures even if success was set earlier, consistently through the MCP Server and `excelcli`.

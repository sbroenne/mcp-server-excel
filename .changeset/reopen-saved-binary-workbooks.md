---
"excelmcp": patch
---

Allow saved `.xlsb` workbooks to be reopened through MCP and the CLI. Previously, Save As supported the binary format but opening its output rejected the file extension. File validation now also accepts `.xlsb` and `.xls` workbooks and checks their openability through a read-only Excel open.

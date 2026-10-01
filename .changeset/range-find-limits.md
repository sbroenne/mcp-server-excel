---
"excelmcp": patch
---

**Accurate range search results** (#949): Find now returns at most 10 matching cells by default, with the exact total, the number returned, and an explicit indication when matches were left out. Both MCP and CLI support a configurable positive match limit; exact counting still searches all matches.

CLI search and replace options now correctly accept JSON objects.

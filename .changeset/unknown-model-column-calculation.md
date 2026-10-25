---
"excelmcp": patch
---

Data Model column results now return `"isCalculated": null` when Excel cannot say whether a column is calculated, rather than incorrectly claiming `false`. Both the MCP server and command-line tool report this as unknown.

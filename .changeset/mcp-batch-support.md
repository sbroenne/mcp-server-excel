---
"excelmcp": minor
---

Add an MCP `batch` tool to run a list of commands in order using the same input format as CLI batch jobs. It continues after errors by default, supports stopping at the first failure, and returns numbered results. It does not roll back after failure; individual commands can save or discard edits.

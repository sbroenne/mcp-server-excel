---
"excelmcp": minor
---

Add an MCP `batch` tool to run a list of commands in order using the same input format as CLI batch jobs. It continues after errors by default, supports stopping at the first failure, and returns numbered results. It does not roll back after failure; individual commands can save or discard edits.

MCP discovery shows the exact batch command and argument names, and file results use `workbook_session_id` consistently. Both entry points stop after timeouts or reported unknown completion, even when continuing after ordinary errors.

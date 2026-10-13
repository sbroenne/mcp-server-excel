---
"excelmcp": patch
---

ExcelMcp now stores a random telemetry ID shared by the CLI and MCP Server.
When Azure telemetry is active at first use, a one-time migration event helps
preserve weekly usage counts. Delete `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id`
and restart the MCP Server to reset the ID.

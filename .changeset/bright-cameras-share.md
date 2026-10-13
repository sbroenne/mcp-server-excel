---
"excelmcp": patch
---

ExcelMcp now stores a random telemetry ID shared by the CLI and MCP Server.
When telemetry is active at first use, a one-time migration event preserves
weekly usage counts. Delete `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id` to reset
the ID.

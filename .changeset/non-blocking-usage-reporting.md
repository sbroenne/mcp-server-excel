---
"excelmcp": patch
---

Keep usage reporting off the critical path for CLI commands and MCP startup, requests, and shutdown. The existing CLI background service sends operation telemetry after short-lived clients exit, without client-side setup or exit waits. CLI reports measure service execution; commands that never reach the service are not reported. Pending reports may still be lost when the service or MCP Server exits.

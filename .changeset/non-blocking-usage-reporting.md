---
"excelmcp": patch
---

Keep usage reporting off the critical path for CLI commands and MCP startup, requests, and shutdown. Slow or blocked reporting no longer adds setup or exit waits; some usage reports may be lost, especially when short-lived commands exit before sending.

---
"excelmcp": patch
---

Stop the Excel MCP server from steadily adding threads and handles while idle. The server still shuts down when its client disconnects or exits unexpectedly.

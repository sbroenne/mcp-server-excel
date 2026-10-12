---
"excelmcp": patch
---

Fix MCP tool calls from Claude Code in Claude Desktop being rejected. The MCP Server now ignores inputs that do not apply to the chosen action when they are empty or set to their advertised default. Real values sent to the wrong action are still rejected with a message naming the input.

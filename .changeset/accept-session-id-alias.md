---
"excelmcp": patch
---

MCP session-bound tools now defensively accept a top-level `sessionId` from client bridges that rewrite the canonical `session_id` argument. The published schema still uses `session_id`, and conflicting or malformed identity values return a privacy-safe input error.

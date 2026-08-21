---
"excelmcp": patch
---

**Canonical public inputs** (#782, #783, #784, #801): CLI, batch JSON, and MCP
now use integer seconds for timeouts, validate the same supported ranges, and
resolve inline-or-file content aliases once at shared dispatch. Manual session
open/create now enforces its documented 10-3600 second operation timeout.

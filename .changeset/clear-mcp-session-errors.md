---
"excelmcp": patch
---

MCP calls missing a required session ID now explain where to supply `session_id` instead of returning a generic tool error. File close errors use the same public parameter name, including when the supplied ID is not a string. This improves diagnosis but does not repair client bridges that drop arguments.

---
"excelmcp": patch
---

**Older Excel formula compatibility** (#750): Formula reads and writes now use the legacy API when Excel does not support modern formulas, in both CLI and MCP. Modern Excel keeps dynamic arrays; older Excel retains its single-value implicit-intersection behavior. Invalid formulas and protected-cell errors are not retried.

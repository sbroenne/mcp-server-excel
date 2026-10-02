---
"excelmcp": patch
---

**Protected range writes by default** (#950): Value/formula writes and content copies now refuse to replace occupied cells, including formulas displaying blank, and report up to 10 conflicting addresses before making any writes. Intentional updates must explicitly select MCP `overwrite_policy: "allow"` or CLI `--overwrite-policy allow`; existing update scripts need this option.

Protected copies check their complete expanded or repeated destination and stop if it cannot be safely inspected. This safeguard does not provide undo, rollback, or protection against interactive Excel edits.

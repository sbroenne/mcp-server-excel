---
"excelmcp": patch
---

**Strict generated action contracts** (#781, #788): CLI direct commands, CLI batch,
and MCP now reject unknown enum values and parameters that do not apply to the
selected action. Power Query load destinations accept the documented
`worksheet`, `data-model`, and `both` aliases without falling back to an
unintended load mode.

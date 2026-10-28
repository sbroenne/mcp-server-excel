---
"excelmcp": patch
---

Shrink the MCP tool list by about 40% so it takes less of the AI's context. Tools no longer advertise result schemas; replies still include the same JSON as text and structured content. Parameter descriptions in MCP and CLI help keep "required for" action lists but drop "valid for" lists, because both entry points already reject parameters an action does not use and name the valid ones. The longest tool descriptions are shorter and no longer repeat parameter details, and action summaries in MCP and CLI help no longer include implementation notes or extra spacing.

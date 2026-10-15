---
"excelmcp": patch
---

Correct Data Model column-type labels in `datamodel list-columns` and `datamodel read-table` for both the CLI and MCP. Numeric columns are no longer mislabeled as small integers, and text, date, and Boolean columns receive accurate labels. Values and loading behavior are unchanged.

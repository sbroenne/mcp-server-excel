---
"excelmcp": patch
---

CLI and MCP now preserve categories for wrapped Excel errors and known VBA and Data Model prerequisites without guessing the cause of unknown failures. Empty macro procedure names are rejected before execution, and Power Query evaluation retains the existing query error categories after cleaning up its temporary objects.

DAX execution errors now identify the failing operation while preserving the underlying Excel error, including when Excel returns only an error code.

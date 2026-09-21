---
"excelmcp": major
---

MCP tool schemas now preserve nullable and mixed-type Excel values using standard SDK schema generation instead of Gemini-specific rewrites. Clients that require Gemini's restricted schema format may no longer accept these tools.

CLI commands now require canonical option names: use `--sheet-name`, `--range-address`, `--session`, `--output`, `--input`, `--quiet`, and `--version` instead of convenience aliases. Update saved commands and automation scripts to use these names.

MCP calls require `session_id` rather than the `sessionId` input fallback. Power Query load modes use `connection-only`, `load-to-table`, `load-to-data-model`, or `load-to-both`; shorthand synonyms are no longer accepted.

Before upgrading the CLI, save and close workbook sessions and stop the old service with the old CLI. Legacy daemon locks and process-tracking formats are no longer migrated.

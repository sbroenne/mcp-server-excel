---
"excelmcp": patch
---

**More reliable MCP tools**: Use the official SDK for registration, schemas,
injected services, and asynchronous calls. Tool failures now set the real MCP
error flag and provide structured results. Every tool now advertises an output
schema generated from its action result contracts, so clients can understand
the returned fields without parsing prose. Unknown, misspelled, wrongly typed,
and action-inapplicable arguments are rejected rather than silently ignored.

Cancellation now reaches workbook startup and reclaims its eventual session
without closing unrelated workbooks. Shutdown cannot publish a late session
after its owner has stopped. Normal shutdown still attempts to save open
workbooks; explicitly closing without saving still discards edits. Expected
client cancellation no longer writes a misleading warning stack trace, while
unexpected handler failures remain logged.

Removed Gemini-specific schema rewriting and generated guide prompts. Shared
guides remain in the skills. Server instructions and skill guidance are now
task-focused, without forced formatting, Table creation, or presentation menus.
Restored detailed parameter documentation in generated tool schemas and corrected
input names, query-loading guidance, and stale skill examples. Bulk-write guidance
preserves the previous calculation mode, and chart feedback no longer requires
screenshots on unavailable desktops. Consent guidance now distinguishes client
confirmation from server-side elicitation, which is not implemented.

Regular PivotTable calculated fields are now recognized as numeric, allowing
them to be added to Values with Sum. Skills distinguish aggregate calculations
from per-row revenue, require complete slicer inputs, and explain recovery when
a failed Power Query load leaves its query behind. The CLI batch example stops
on failure and explicitly discards only the failed job's own unsaved changes.
Both skills now include native examples, a complete guide index, and less
repeated guidance; the CLI command catalog is split into smaller linked pages.

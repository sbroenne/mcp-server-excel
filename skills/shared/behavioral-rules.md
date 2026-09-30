# Working safely with Excel

Discover the intended workbook, sheets, and objects before changing them. Reuse a
matching session, not an arbitrary open file. Ask only for unresolved targets or
destructive decisions. Never invent a private path. Reading does not require
formatting, Tables, charts, or PivotTables.

## Sessions and failures

- Use the returned session ID on every follow-up. CLI and MCP sessions are
  separate; their IDs cannot be transferred.
- Close only when authorized, operations have finished, and the session listing
  reports `canClose: true`. Confirm before closing a visible window unless
  already authorized. Keep the workbook open when requested.
- Explicit close discards unsaved edits unless saving is requested. Normal
  service shutdown attempts to save remaining sessions; leaving a failed job
  open is not a rollback. Crashes and forced cleanup can lose changes.
- Cancellation is not undo. After failure, inspect the surviving session and
  affected objects before retrying. A failed operation can partly apply.
- Save only the intended successful result. For a session opened exclusively for
  a job, close without saving after failure. Do not discard another user's
  existing session or earlier unsaved work.
- Report what actually succeeded, the saved file when relevant, and any remaining
  failure. Do not present an attempted action as a completed result.

Use the file test operation when access or protection is uncertain. It reports
`canOpen`, `isIrmProtected`, `willOpenReadOnly`, and `requiresVisibleSession`.
Ordinary files are briefly opened read-only for this check. IRM/AIP workbooks
require interactive Excel authentication; do not work around protection.

## Changes and formatting

Make targeted writes and prefer resize, rename, refresh, or update over rebuilding
objects. Deleting objects can break formulas, relationships, measures, and charts.
Check their dependencies first.

Use the owning object's style system: Table styles for Tables, chart styles for
charts, and range formatting for plain cells. Do not style PivotTable cells with
range formatting; refresh overwrites it. Combine visual properties in one call;
use shared multi-range formatting for repeated styles on one sheet.

Use US number-format codes; Excel displays them in the user's locale. Preserve
existing formats and fixed layouts unless a change is requested. See
[ranges and formatting](range.md) for examples.

For costly bulk writes, get the current calculation mode with `get-mode`, switch
to manual, calculate after writing, and **restore the prior mode** in `finally`.
Reads and operations needing intermediate results do not need manual mode.

## Inputs and errors

Use only the selected action's parameters. Supply either inline content or a
readable source file, never both. Timeouts are integer seconds, not duration
strings. Read the action's actual limits; session timeouts and data refresh
timeouts are different.

Read `errorMessage`, `errorCategory`, and `suggestedNextActions` when present.
Correct input, prerequisites, or access before retrying. Missing Data Model tables
and missing MSOLAP installation are different failures. A generic Excel error
does not establish a VBA trust problem or invalid query. Never change security
settings automatically.

Remote M/DAX formatting is opt-in and sends code to an external service. Obtain
explicit consent first. Follow [Power Query](powerquery.md) and
[Data Model](datamodel.md) guidance rather than repeating writes blindly.

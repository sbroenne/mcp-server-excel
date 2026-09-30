# Gotchas and Known Limits

## Formatting and Data Structures

- PivotTable cell formatting may be replaced on refresh. This API does not expose
  PivotTable visual styles; there is no `pivottable set-style` action. Use
  field number-format operations and calculated/layout operations.
- Excel Table styling belongs to Table styles, not plain-range visual formatting.
- Worksheet Tables and Data Model tables are separate. After changing a source
  Table, refresh the Data Model before relying on DAX results.
- Data Model metadata is limited to what Excel exposes. Inspect the actual
  list/read results rather than assuming an object is missing because it is hidden.

## Sessions, Visibility, and Recovery

Use `file list` before opening a workbook. Reuse the matching session; do not
open the same workbook in a second Excel instance.

Each session owns an Excel instance and its COM work runs on that session's
dedicated thread. Operations in one session are serialized. Clients do not
manage COM threads. `window show/hide` affects the selected session's Excel
application, not every session.

Explicit close defaults to discarding edits. Set `save: true` to keep them and
wait for `canClose: true`. Confirm before closing a visible window unless already
authorized; leave it open when requested. Normal shutdown attempts to save
remaining sessions, but crashes and forced cleanup may lose changes.

Cancellation is not undo. Inspect the current sessions and affected workbook
state before retrying a write, recreating an object, or reopening a file.

## Timeouts and Refresh

Session open/create accepts integer timeout seconds from 10 through 3600.
Power Query refresh/refresh-all accepts 0 through 2147483; omitted or zero uses
the 30-minute data-operation default. Its data-operation timeout owns that
refresh, rather than layering another session wait over it.

Other generated timeout actions accept 1 through 2147483 seconds. Use the
installed action schema for the supported range and default.

External source changes are not automatically reflected in loaded query data.
Refresh the relevant query when the task requires current data.

Connection-string keys and values follow the selected provider's rules; they
are not universally case-sensitive. Use `connection test` to check a connection,
and read refresh errors. Never expose credentials or full connection strings
in summaries or diagnostic artifacts.

## Calculation and Python

Formula results depend on the current calculation mode and state; writing a
formula does not universally produce zero or require manual calculation.
Read the calculation mode, calculate when needed, and restore the prior
mode after any temporary change. Read formula text with `range get-formulas`.

```mcp
calculation_mode(action: 'get-mode', session_id: sessionId)
calculation_mode(action: 'calculate', session_id: sessionId, scope: 'workbook')
```

```cli
excelcli -q calculationmode get-mode --session $sessionId
excelcli -q calculationmode calculate --session $sessionId --scope workbook
```

`pythoninexcel` executes in Microsoft's cloud, not local Python. It requires a
licensed Microsoft 365 account with Python in Excel and internet access.
`#NAME?` means the feature is unavailable; do not treat it as pending work.
For pending cloud results, follow `get-result` status and error guidance.
Do not assume every connection or policy error is transient.

CLI `--max-wait-seconds` must be at least 1 and shorter than the session's
operation timeout. Cloud startup may take several minutes.

## Addresses and Dates

Use A1 addresses such as `A1:D10` or a valid Excel named range. Pass worksheet
names as strings in the worksheet-name input; quotes are part of the JSON/call syntax, not
part of the worksheet name. A sheet-qualified Excel reference with spaces uses
`'Sales Data'!A1:D10`; backticks are not Excel reference quoting.

Date cells may be returned as numeric serials. Do not treat them as Unix
timestamps or Python ordinals. Check the workbook's 1900/1904 date system and
the result's meaning before converting outside Excel; Excel's 1900 calendar
also has a historical leap-year exception. Prefer Excel date formatting when
the task only requires readable dates.

## Visual Verification

Screenshots require an interactive, unlocked Windows desktop. They cannot be
required for every unattended job. Inspect chart bounds and overlap warnings
when a screenshot is unavailable, and say visual verification was not performed.

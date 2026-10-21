# Workbook Lifecycle

Inspect the intended workbook's saved/read-only state, properties, and external
dependencies before changing its format or publishing output. Session lifecycle
is separate; see [session and saving guidance](behavioral-rules.md#sessions-and-failures).

Current commands and supported inputs come from CLI help or MCP tool descriptions.

## Opening a SharePoint workbook

Open accepts a direct SharePoint or OneDrive for Business HTTPS file URL ending
in `.xlsx`, `.xlsm`, `.xlsb`, or `.xls`. An optional `?web=1` or `?web=0` is
removed before opening and session matching. Spaces and their percent-encoded
forms identify the same workbook. Folder URLs, browser pages such as `Doc.aspx`,
sharing links, other query parameters, arbitrary websites, and OneDrive personal
consumer links are not supported.

For an unsupported link, open the file in desktop Excel and use
**File > Info > Copy Path** to obtain its direct workbook URL, or supply the
path to an existing local workbook. ExcelMcp does not register a Microsoft app
or resolve opaque sharing links.

Before publishing a session for a `.xlsx`, `.xlsm`, `.xlsb`, or `.xls` location,
ExcelMcp checks Excel's actual file format. A sign-in page, browser response, or
text import is not accepted as a successful workbook open, even if Excel can
display it as a worksheet.

Use MCP `file(action: 'open', file_path: '<direct-url>', show: true)` or
`excelcli session open "<direct-url>" --show`. The visible session is required
so Office sign-in and rights-management prompts remain accessible. Excel uses
the signed-in Office account; ExcelMcp neither supplies credentials nor bypasses
IRM/AIP restrictions. Inspect `workbook_read` action `get-info` or
`excelcli workbook get-info` for `readOnly` and the live `autoSaveOn` status before
editing. When `autoSaveOn` is true, Excel can persist changes without an explicit
save.

For `workbook_read get-info`, Excel versions without the AutoSave property
report `autoSaveOn: false`; there is no AutoSave feature to disable on those
versions. Discovery and attachment report AutoSave as unavailable instead.
Only known unavailable-member errors are treated this way. Other AutoSave read
or write failures remain errors.

`file_read` action `test` and `excelcli session test` report that URL validation
requires interactive opening. They do not download the workbook or probe a local
file. Returned `exists: false`, zero size, and `isIrmProtected: false` are unknown
remote metadata, not evidence of a missing or unprotected workbook.

Cloud AutoSave is disabled in editable remote sessions opened by ExcelMcp.
Explicit save/close behavior therefore matches local workbooks: close without
saving discards unsaved session edits; close with saving writes through Excel to
the original SharePoint location. Normal Service shutdown still attempts to save
workbooks it opened.

If the remote workbook is already open, use MCP `file_read` action `discover`
or CLI `excelcli session discover`, then attach using MCP `file` action `attach`
with `open_workbook_id` or CLI `excelcli session attach --workbook-id <ID>`.
Discovery does not reopen workbooks and reports `complete` plus any errors.
Newly attached workbooks remain open on service shutdown. If a workbook is
already connected to a session, attachment returns that existing session and
its established shutdown behavior applies. The CLI tray's Exit prompt applies
only to workbooks opened or created by ExcelMcp; either save choice leaves
attached workbooks open without saving them. Explicit Close Session and Close
All Sessions still close the selected sessions, including attachments.
Attachment preserves the existing
AutoSave setting and reports it; when AutoSave is enabled, Excel may save later
edits automatically, even without an explicit save.
For already-open remote workbooks, discovery reports the HTTPS location without
URL credentials, query parameters, or fragments. This does not change which URLs
are allowed when opening a workbook through ExcelMcp.
Discovery also lists workbooks in Excel processes already used by this host.
`sessionId` is included for workbooks already connected to a session; attaching
one returns that existing session. Other workbooks in the same process can be
attached separately. Commands sharing an Excel process run one at a time,
including restoration of temporary Excel settings. Workbooks in separate
processes can still be used concurrently.
Showing or hiding an Excel process through one session updates the confirmed
visibility reported for all its local connections, including owned and attached
sessions. Listing that state does not wait for a running Excel command.
Older callback observations and deferred timeout reveals cannot undo a newer
show or hide requested through another connection to that process.
Workbook IDs identify an open Excel workbook instance; a
closed and reopened workbook receives a new ID, and a complete discovery retires
any stale attached session for the closed instance.
A scan can retire only connections that existed before it began; it cannot
remove a workbook connected by a newer scan or undo a newer Save As location.
If a connection or tracked location changes during a scan, its results can be
incomplete with a request to discover again.
Discovery cannot replace a tracked location while Save As is still pending.
An overlapping reconnect request fails with a request to discover again if
Save As is pending or has changed the connection since that request began.
This also applies if Save As changes the path and then changes it back.
Ordinary commands that observe a rename in Excel also invalidate older location
results; a later scan cannot restore the old name or release the newer path.
Saving an owned workbook under a new name in Excel does not create a second
session: discovery reconciles its saved location and attachment returns its
original session.
Overlapping Save As protection also covers workbooks opened by ExcelMcp:
older discovery or reconnect results cannot replace the new location or
create another session for that workbook. Discovery rechecks the live workbook
before applying an Excel-side rename.
Reattaching an already-connected workbook also updates its tracked name and
location after Excel Save As, without requiring another discovery first.
Save As through ExcelMcp updates the attached workbook's listed name and path
immediately; resolving a location conflict does not release another session's
path reservation.
Discovery releases cached connections and selections only when the exact Excel
process is confirmed to have exited, not merely because its window is missing.
After connection cleanup fails, discovery or shutdown can release that connection
once both its exact Excel process exit and its execution thread's exit are confirmed.

On shutdown, ExcelMcp closes only workbooks it opened. If another workbook
remains in that Excel process, ExcelMcp leaves the process running. ExcelMcp
checks for remaining workbooks before every attempt to quit, including retries,
and restores the alert setting it changed when leaving its Excel process open.
Disconnecting from user-started Excel does not change that application's alert
setting. ExcelMcp does not automatically terminate Excel when its host exits unexpectedly, so
Excel can preserve user workbook state.
If an attached workbook still has work running, shutdown refuses to disconnect
it and keeps its session and operation status available until that work ends.
After a wait expires, inspecting a healthy attached workbook cannot clear
recovery protection while another workbook in the same connection is refreshing.
Readiness must be confirmed across that connection before ordinary commands resume.

A successful explicit save confirms Excel's saved state, not independent
server-side upload completion.
Save As, Save Copy As, export, and create continue to require Windows output paths.
Save Copy As preserves the source format. For a never-saved workbook, Excel's
actual workbook format determines the required output extension; copying does
not give the source a saved path or save its pending edits.

## Metadata and document properties

Use built-in properties for existing document metadata and custom properties
for workbook-specific information. Custom properties can be removed; built-in
ones are maintained by Excel. Avoid putting private paths, credentials, or
connection strings in published properties.

MCP `workbook_read` actions `list-document-properties` / `get-document-property`
and `excelcli workbook list-document-properties` / `get-document-property`
return `valueAvailable` and `valueError` for each property. Excel raises an error
for some undefined built-in values, but does not distinguish those from a failed
`Value` getter. Such a property retains its name, scope, and type, with
`valueAvailable: false`, a null value, and the native error in `valueError`.
This is uncertainty, not proof of absence. Successfully read values, including
empty strings, have `valueAvailable: true` and no value error. Other unexpected
native errors still fail the operation.

Use MCP `workbook_read` action `inspect` or `excelcli workbook inspect` for a
bounded workbook overview. It returns worksheet visibility and used-range
geometry, table locations, and visible user-defined names, with per-section
counts and omitted-item counts. `sheet_name` narrows worksheet and table
metadata; the named-range section remains workbook-wide. Hidden worksheets are
reported, not skipped. The operation reads through Excel and does not change the
workbook or its view.

Cell previews are optional. Set `include_preview=true` and provide `sheet_name`;
optionally set `range_address` to limit the source range. Without it, the
preview starts at the top-left of Excel's UsedRange. The preview reads at most
10 rows by 10 columns, and reports omitted rows and columns. `max_cell_characters`
and `max_preview_characters` further bound text in returned values and formulas.
Row and column limits are applied before reading cell contents; text limits are
applied to the response. For example:

```text
MCP: workbook_read(action: 'inspect', workbook_session_id: sessionId,
     sheet_name: 'Summary', include_preview: true, range_address: 'A1:F100',
     max_preview_rows: 5, max_preview_columns: 6)
CLI: excelcli -q workbook inspect --session <session-id> --sheet-name Summary
     --include-preview true --range-address A1:F100 --max-preview-rows 5
     --max-preview-columns 6
```

Excel reports an empty worksheet's UsedRange as `$A$1`. Its preview contains a
null value and an empty formula; the returned address does not mean the cell is
populated.

If Excel cannot read a protected or otherwise inaccessible range, the operation
reports an error rather than presenting it as empty.

## Document Inspector

Use MCP `workbook` action `inspect-document` or
`excelcli workbook inspect-document --session <session-id>` to run Excel's
available Document Inspector modules. Unlike `inspect`, which summarizes
workbook contents, this checks categories of hidden information.

Each returned inspector includes its one-based `index`, exact localized `name`,
native `statusCode`, stable `status` (`ok`, `issue-found`, `error`, or `unknown`),
and localized `results` text. Findings do not mean inspection failed. Native
errors, unknown status codes, or no available modules are reported as failures.
Completed checks remain in the result when another module reports an error.

Inspection does not request removals or save the workbook. However, Excel's
native inspector modules can mark a saved workbook as having unsaved changes,
even when they report no findings. The response includes `savedBefore` and
`savedAfter`; ExcelMcp does not reset that flag to hide the change. For this
reason, MCP exposes inspection through `workbook`, not the read-only
`workbook_read` tool. Inspection still supports read-only workbooks.

This is not a full privacy audit or a guarantee that sensitive information is
absent. Excel exposes only some built-in modules and installed custom inspectors.
Comments/revisions and document properties are not included in this collection.
Module availability and names depend on the installed Excel version and language;
do not assume a particular category always has the same index.

To request removal, use `fix-document-inspector` with both the returned index and
the exact name, including any ampersands:

```text
MCP: workbook(action: 'fix-document-inspector', workbook_session_id: sessionId,
     inspector_index: selected.index, inspector_name: selected.name)
CLI: excelcli workbook fix-document-inspector --session <session-id>
     --inspector-index <returned-index> --inspector-name "<returned-name>"
```

The index and name must still match the live collection. Read-only or
structure/window-protected workbooks are rejected before removal.

**Removal is category-wide, not item-by-item.** For example, fixing hidden
worksheets can delete the hidden worksheets and all their data and affect
formulas referring to them. Request it only when that entire removal is intended.
There is no tool-level undo. The action does not save automatically, although
existing Excel AutoSave can persist changes independently. Save explicitly to
keep changes.

The response includes the native `fix` outcome and a new `inspection` of the
selected module. Success requires a successful fix and no remaining findings.
An error does not promise rollback: changes may already have occurred. Inspect
the workbook's current state before deciding whether another action is needed.
The server does not display its own confirmation dialog.

## Save and publish

Opening or creating a workbook requires a confirmed Excel process ID and start
time for safe readiness checks and cleanup. Startup retries a temporary capture
failure. If the identity remains unavailable, it reports an error before opening
or creating the workbook rather than returning a session that cannot save or close.

Inspect `readOnly` with MCP `workbook_read` action `get-info` or
`excelcli workbook get-info` before editing. Protected workbooks require visible
authentication; Excel decides editing rights. Do not change protection to work
around a genuine permission restriction.

Workbook-changing actions reject read-only access before editing. Inspection,
window controls, calculation other than workbook precision changes, and authorized
Save As/copy/export remain available. MCP `calculation_mode` action `set-precision`
and `excelcli calculation set-precision` require editing rights.
Excel still enforces permissions on outputs.

Writes change the open workbook, not necessarily the saved file. Saving a
read-only workbook or a save cancelled by Excel returns an error. A failed save
leaves the session open and any unsaved changes available for inspection; do not
assume they were persisted or automatically discard them.

Successful saving confirms Excel's saved state, not completion of OneDrive
synchronization or upload to SharePoint.

Saving, Save As, and explicit closing check Excel's live refresh readiness.
A running refresh, open modal dialog, busy Excel instance, or unconfirmed status returns a `Busy`
error without saving or discarding edits. The session remains open. Wait for
Excel to finish, check `canClose` through MCP `file_read` action `list` or
`excelcli session list`, inspect the refreshed values, then retry. A zero
`activeOperations` count alone is not proof that Excel is idle.

Readiness inspection waits at most 200 milliseconds, even when its Excel scan
has started. An unfinished scan remains owned and reports an unconfirmed state;
later listings reuse its result. This does not save or close the workbook.

Service shutdown also refuses to discard edits when readiness blocks saving.
Normal close checks readiness again on Excel's STA at the shutdown transition.
If a refresh starts after the initial check, close reports `Busy` without
cancelling the work queue or retiring the session, so it can be used and retried.
When `save=true`, saving may already have completed before this late refusal;
that save is not rolled back.
The blocked workbook stays in its session and the service stays running so you
can respond to Excel's prompt or wait for the refresh, then retry
`excelcli service stop`. Other ready sessions may already have been saved and
closed before a refusal. Product `excelcli service stop` waits for the shutdown
reply while ready workbooks save and close one at a time. You can cancel this
wait; Excel's own save and shutdown time limits still apply. It does not force-stop Excel
after any failed shutdown response or an expired shutdown wait.
Simultaneous shutdown requests cannot bypass a refusal: saving and the decision
to stop the service are serialized across all service shutdown entry points.
This does not protect against externally killing the service or Excel.

Development builds use a separate force-stop policy: they stop only this
worktree's CLI background services without requesting graceful shutdown or
saving workbooks. Unsaved development-session work may be lost. They do not
terminate Excel, MCP, foreground CLI commands, or other worktrees' services.

The same session listing exposes `excelState`, `operationPending`, and
`blockingReason`. A confirmed blocking dialog reports `dialogOpen` and
automatically reveals the session's Excel window. Ask the user to check the
window and respond there. Readable dialog titles and message labels are included
automatically in `blockingMessage`; unreadable or partial text is marked.
The operation remains pending; the server never identifies the dialog type,
reads input values, answers, or dismisses the dialog. After responding,
check readiness and refreshed data before saving.

If an executing operation reaches the session's operation timeout while a dialog
is detected, its error reports that a dialog was open instead of only suggesting
a longer timeout. Readable dialog text is included automatically. Timeout recovery retains the
session and unsaved workbook and requests visibility for its owned Excel window.
It does not save, close, terminate, or roll back the operation. Session listing
checks window ownership and reads available dialog labels without queuing Excel COM work.
The operation may still be running. CLI `session list` and MCP `file_read`
action `list` include `operation` with `state`, `command`, `elapsedSeconds`,
`waitExpired`, `executionPending`, `errorMessage`, and `windowReveal`.
Reveal outcomes distinguish `alreadyVisible`, `shown`, `showRequested`, and
`unavailable`; a request is not confirmation that Excel became visible.
Late completion or failure remains visible until new work starts. New work,
saving, and closing are refused while execution or readiness remains unresolved.
Once the original callback and cleanup have returned and Excel is ready, the
same session can be reused. These rules also apply to startup waits; an expired
open, create, or test wait reports the retained session ID.

Save the current workbook in place with `excelcli session save -s <sessionId>`
or MCP `file(action: 'save', workbook_session_id: sessionId)`. Saving after an
expired wait remains a caller decision, not automatic recovery.

Power BI/MSOLAP (OLAP) connections always refresh synchronously. Their unsupported
background setting is reported as false; enabling it is rejected before changing
other properties. An empty last-refresh date is reported as unknown, not as proof
of a completed query.

Choose the output according to the task. Saving under a new name changes the
active workbook's path; saving a same-format copy leaves the active workbook
unchanged. PDF/XPS output is a published view, not an editable workbook.

Saved `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` workbooks can be reopened with
`file(action: 'open', file_path: ...)` (MCP) or `excelcli session open <FILE>` (CLI).

Replacing an output needs authorization. A format conversion can remove
features: saving a macro-enabled workbook as `.xlsx` removes its VBA content.
Inspect the result and active path before continuing work.

For unattended exports, avoid opening viewers or modal previews. Configure
the [print layout](worksheet.md#styling-and-outlines) before publishing.

## External Excel links

Discover the actual linked sources before refreshing them. A refresh can change
calculated results; it is not merely inspecting the workbook.

Breaking a link permanently replaces linked formulas with their current values.
Do it only when that change is intended, and inspect the resulting values before
saving. There is no tool-level undo. Closing without saving loses all earlier
unsaved work too.

Printing and print preview are not exposed: a default printer can produce
physical output, and modal preview can block unattended sessions.

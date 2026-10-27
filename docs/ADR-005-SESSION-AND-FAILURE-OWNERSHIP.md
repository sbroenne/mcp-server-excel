# ADR-005: Own Excel lifetime and failures at session boundaries

**Status:** Current

## Context and decision

A session owns its workbook context. A private worker owns execution for the
Excel application containing that workbook, including attached neighbors.
Independent applications may run independently; workbooks sharing an application
must not execute concurrently. ComInterop owns the execution queue and cleanup,
while the worker's Service supplies routing and structured failure responses.
The host controller owns recorded session routing, not Excel COM objects.

An expired wait is not proof that Excel stopped or failed. The owning session,
workbook, path reservations, and exact process identity remain retained. The
affected Excel window is revealed through native window operations when possible,
without waiting behind the outstanding COM call. CLI and MCP return the expired
wait and basic operation status to their calling agent; its recovery decision is
outside the server's scope. Neither deadline expiry nor caller cancellation
authorizes saving, discarding, closing, terminating, or rolling back workbook work.
This also applies to startup that finishes after its caller stops waiting.
Core lets execution failures reach these owners rather than translating a COM
failure into absent data. Optional property discovery tolerates only an exact
known unavailable-member error; unexpected errors retain their cause and
operation context at the failure-transport boundary.

Normal host shutdown does not authorize saving or discarding pending edits.
Saved workbooks opened by the host close; unsaved workbooks are handed back open
and visible with automation settings restored. Attached workbooks stay untouched.
Opening and attaching preserve Excel AutoSave; an explicit per-workbook operation
controls it. AutoSave can persist edits independently, including before shutdown.

Shutdown follows these rules:

- `SessionManager` decides what happens to each Excel process. Discovery
  connections and the tray never quit Excel themselves; a discovery connection
  quits Excel only when the host owns no session in that process.
- A workbook belongs to the host only as the exact COM object the host opened.
  A workbook the user closed in Excel is forgotten without saving; a file the
  user reopens is the user's workbook. Saving a workbook the user already closed
  fails and reports that nothing was saved.
- Checks never change workbook state, and a timed-out check means unknown, not
  closed. The saved-state check and close run in the same Excel step with Excel
  prompts enabled, so edits made after the check are protected by Excel itself.
- Excel quits only when the host started it and no workbook remains open.
- Each Excel process has its own bounded shutdown wait. New requests are refused
  once shutdown starts. MCP shutdown never prompts; a tray answer applies only
  to the workbooks it listed.
- When shutdown cannot confirm a process's state (busy, not responding, or
  failed), the host still exits. That Excel is left exactly as it is, revealed
  when possible, and reported by process number and start time. "Service stopped"
  never implies that Excel exited.
- A stopped controller can leave a worker holding pending Excel execution.
  Its ownership does not end merely because the controller stopped waiting.
  A replacement controller may claim that same worker only after the previous
  controller's exact process is gone and the build matches. Discovery and
  attachment are explicit. If the worker exited, discovery reports the unknown
  prior outcome before reconnecting to Excel for inspection. Nothing repeats
  the old edit or terminates the retained application.
  A surviving worker reveals retained workbooks after controller loss. A new
  build cannot claim an older worker; those workbooks remain available for
  direct inspection in Excel without saving or closing them automatically.

## Reasons and tradeoffs

Opening Excel per operation would avoid persistent session state but lose
in-memory continuity and repeatedly pay startup costs. Concurrent calls on
the same Excel object would conflict with its threading and state model.

Stopping a caller's wait does not necessarily stop Excel's current COM call.
Continuing to use that batch immediately would treat uncertain state as healthy.
New work is refused until the original callback and its cleanup have returned
and a bounded readiness check confirms Excel is ready. The same session can then
be reused. Basic status retains late success or failure; readiness alone never
establishes the original operation's result. Cancellation is not a transaction
rollback. Preserving uncertain work can keep Excel and its owner alive longer,
but avoids silently losing unsaved changes.

A failed opening remains discoverable until explicitly closed, even when its
Excel process has already exited; status polling must not erase its late error.
Status polling does not wait for workbook-location bookkeeping or clean up dead
sessions. Completed Save As paths are reconciled by subsequent workbook access,
not by status polling, and their reservations remain held until reconciliation.

Status checks have a short wait budget and may report unknown while their owned
readiness scan finishes. Normal save and close verify workbook readiness on the
Excel thread within the operation's budget; a slow scan is not itself a failed
operation. An expired operation still requires a completed recovery scan before
new work is accepted. Status visibility reflects any reveal applied by that scan.
Session operations check for confirmed native dialogs before workbook write
validation, because the write guard already accesses Excel COM. Window show/hide
remain available. Save and Save As check again inside their callback for races.
This does not treat an information-protection banner as a modal dialog or
grant permission to answer either.
Validation that opens a workbook but refuses close before shutdown commits
retains a usable session, not a failed-teardown quarantine. After readiness
returns, the caller can inspect and explicitly close that session.

Automatic CLI startup also preserves Excel instances left alive after their
previous service exits. It carries their exact process identities forward
separately from workbooks owned by the new service, including across repeated
restarts. They do not block startup and are not attached automatically.
Service status and stop report these still-running Excel copies with their
process numbers, start times, and unknown workbook state. Discovery and explicit
attachment let the caller inspect the intended workbook without reopening it.
An expired shutdown wait or an unconfirmed shutdown response does not authorize
forced termination of a running service or its workbooks.
Explicit service stop attempts bounded pipe shutdown even without a daemon mutex.
When no service answers and no service mutex is held, stop reports that the
service is not running and leaves recorded Excel copies untouched. Only obsolete
records with no confirmed live processes are removed. Damaged or unreadable
records remain errors, not evidence that Excel has exited.

Central ownership makes it possible to report exactly which Excel copies remain
without terminating unrelated user workbooks. The caller decides whether any
further recovery action is appropriate.

Implicit saving at disconnect commits changes without a separate request;
implicit discard loses the user's choices. A visible handover preserves those
choices at the cost of leaving Excel running. Unattended callers must explicitly
save and close their completed work. Handover must survive host exit and allow
later discovery/attachment.

Refusing to stop while Excel is busy would keep a service running that the user
asked to stop, with no way forward when Excel never responds. Exiting and
reporting the exact Excel process leaves the decision with the user or calling
agent, which may close that Excel itself.

## Implementation and guidance

- [Batch execution](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Session ownership](../src/ExcelMcp.ComInterop/Session/SessionManager.cs)
- [Shutdown](../src/ExcelMcp.ComInterop/Session/ExcelShutdownService.cs)
- [Cancellation tests](../tests/ExcelMcp.McpServer.Tests/Unit/ServiceBridgeCancellationTests.cs)
- [Lifecycle semantics](../CONTEXT.md#runtime-relationships) and [COM instructions](../src/ExcelMcp.ComInterop/AGENTS.md)

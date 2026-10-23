# ADR-005: Own Excel lifetime and failures at session boundaries

**Status:** Current

## Context and decision

A session owns an Excel process and its workbook context. Its operations run
in order on one dedicated Excel-compatible thread; independent sessions may
run independently. ComInterop owns the execution queue and cleanup, while
the Service supplies shared session routing and structured failure responses.

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
and visible with automation settings restored and their process protected from
cleanup. Attached workbooks stay untouched. Opening and attaching preserve Excel
AutoSave; an explicit per-workbook operation controls it. AutoSave can persist
edits independently, including before shutdown.

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
Save and Save As check for native dialogs before batch execution, because its
write guard already accesses Excel COM. They check again inside the callback
for races.
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
later discovery/attachment, not remain an orphan eligible for process cleanup.

## Implementation and guidance

- [Batch execution](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Session ownership](../src/ExcelMcp.ComInterop/Session/SessionManager.cs)
- [Shutdown](../src/ExcelMcp.ComInterop/Session/ExcelShutdownService.cs)
- [Cancellation tests](../tests/ExcelMcp.McpServer.Tests/Unit/ServiceBridgeCancellationTests.cs)
- [Lifecycle semantics](../CONTEXT.md#runtime-relationships) and [COM instructions](../src/ExcelMcp.ComInterop/AGENTS.md)

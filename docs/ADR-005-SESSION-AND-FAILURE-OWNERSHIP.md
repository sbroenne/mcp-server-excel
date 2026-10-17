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

Status checks have a short wait budget and may report unknown while their owned
readiness scan finishes. Normal save and close verify workbook readiness on the
Excel thread within the operation's budget; a slow scan is not itself a failed
operation. An expired operation still requires a completed recovery scan before
new work is accepted. Status visibility reflects any reveal applied by that scan.

Automatic CLI startup also preserves Excel instances left alive after their
previous service exits: it refuses to replace their recorded ownership while they
remain live. An expired shutdown wait or an unconfirmed shutdown response does
not authorize forced termination of a running service or its workbooks.

Central ownership also makes targeted cleanup possible without terminating
unrelated user workbooks. This carries explicit lifetime and recovery complexity
rather than hiding it in individual commands.

## Implementation and guidance

- [Batch execution](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Session ownership](../src/ExcelMcp.ComInterop/Session/SessionManager.cs)
- [Shutdown](../src/ExcelMcp.ComInterop/Session/ExcelShutdownService.cs)
- [Cancellation tests](../tests/ExcelMcp.McpServer.Tests/Unit/ServiceBridgeCancellationTests.cs)
- [Lifecycle semantics](../CONTEXT.md#runtime-relationships) and [COM instructions](../src/ExcelMcp.ComInterop/AGENTS.md)

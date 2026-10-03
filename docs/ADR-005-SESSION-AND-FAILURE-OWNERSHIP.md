# ADR-005: Own Excel lifetime and failures at session boundaries

**Status:** Current

## Context and decision

A session owns an Excel process and its workbook context. Its operations run
in order on one dedicated Excel-compatible thread; independent sessions may
run independently. ComInterop owns the execution queue and cleanup, while
the Service supplies shared session routing and structured failure responses.

A timed-out executing batch is no longer usable. The MCP bridge ties
cancellation cleanup to the Service instance and session that owned the
request, including a session that finishes opening after cancellation.

## Reasons and tradeoffs

Opening Excel per operation would avoid persistent session state but lose
in-memory continuity and repeatedly pay startup costs. Concurrent calls on
the same Excel object would conflict with its threading and state model.

Stopping a caller's wait does not necessarily stop Excel's current COM call.
Continuing to use that batch would treat uncertain state as healthy. Closing
the affected session favors a clear failure over misleading success, but can
lose unsaved work. Cancellation is not a transaction rollback.

Central ownership also makes targeted cleanup possible without terminating
unrelated user workbooks. This carries explicit lifetime and recovery complexity
rather than hiding it in individual commands.

## Implementation and guidance

- [Batch execution](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Session ownership](../src/ExcelMcp.ComInterop/Session/SessionManager.cs)
- [Shutdown](../src/ExcelMcp.ComInterop/Session/ExcelShutdownService.cs)
- [Cancellation tests](../tests/ExcelMcp.McpServer.Tests/Unit/ServiceBridgeCancellationTests.cs)
- [Lifecycle semantics](../CONTEXT.md#runtime-relationships) and [COM instructions](agents/rules/excel-com-interop.md)

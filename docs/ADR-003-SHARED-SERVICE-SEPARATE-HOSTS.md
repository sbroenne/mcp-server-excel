# ADR-003: Share Excel behavior, not live entry-point sessions

**Status:** Current

## Context and decision

The MCP Server and `excelcli` are equal product entry points backed by the same
Service and Core operations. Each host owns an `ExcelController`; CLI invocations
connect to their persistent controller over named pipes, while the MCP bridge
calls its own controller directly. Neither entry point depends on the other's
host or live sessions.

Excel COM objects and `ExcelMcpService` live in private worker processes, one
execution owner per Excel application. Workers launch from the current product
executable before normal CLI/MCP startup. A controller records sessions and
process identities, routes work, and answers basic status without calling Excel.
The tray reads those records, not Excel. The shared listener owns accepting and
draining named-pipe connections. Core does not own either transport.

## Reasons and tradeoffs

A persistent daemon lets short CLI invocations operate on an already-open
workbook. MCP already has a long-lived host, so it owns its controller rather
than depending on the CLI daemon.

A blocked Excel call can also prevent the thread waiting for it from finishing.
Keeping Excel execution outside the controller protects status, stopping, and
tray interaction from that wait. This does not make an arbitrary Excel call
interruptible or establish the cause of the reported hangs in #1249.

The worker connection uses the existing StreamJsonRpc and protected Windows
named pipes. Standard .NET hosting and process-exit monitoring own the private
process lifetime. The application lease still prevents competing Excel owners.
The additional process boundary costs startup time and requires explicit
handling of disconnected callers, uncertain results, and ownership handover.

Separate implementations of Excel behavior would drift in defaults and
validation. A shared Service keeps those responsibilities in one place while
allowing different connection lifetimes.

Separate hosts mean a session opened through one entry point is not available
through the other. Sharing command behavior is not a promise of identical wire
names, output envelopes, or session identifiers across hosts.

## Implementation and guidance

- [Service](../src/ExcelMcp.Service/ExcelMcpService.cs)
- [Controller](../src/ExcelMcp.Service/ExcelController.cs)
- [Private workers](../src/ExcelMcp.Service/Workers/ExcelWorkerHost.cs)
- [CLI daemon host](../src/ExcelMcp.Service/Rpc/DaemonHost.cs)
- [MCP bridge](../src/ExcelMcp.McpServer/ServiceBridge/ServiceBridge.cs)
- [System map](../CONTEXT.md) and [runtime instructions](../src/AGENTS.md)

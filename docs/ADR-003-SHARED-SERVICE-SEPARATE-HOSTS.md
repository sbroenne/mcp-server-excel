# ADR-003: Share Excel behavior, not live entry-point sessions

**Status:** Current

## Context and decision

The MCP Server and `excelcli` are equal product entry points backed by the same
Service and Core operations. MCP owns an in-process Service through a
host-owned bridge. CLI invocations connect to a persistent named-pipe daemon.
Their live session sets are separate.

The Service owns routing and sessions; the daemon host owns accepting and
draining pipe connections and idle shutdown. Core does not own either transport.

## Reasons and tradeoffs

A persistent daemon lets short CLI invocations operate on an already-open
workbook. MCP already has a long-lived host, so routing it through another
daemon would add a process/transport dependency without serving that need.

Separate implementations of Excel behavior would drift in defaults and
validation. A shared Service keeps those responsibilities in one place while
allowing different connection lifetimes.

Separate hosts mean a session opened through one entry point is not available
through the other. Sharing command behavior is not a promise of identical wire
names, output envelopes, or session identifiers across hosts.

## Implementation and guidance

- [Service](../src/ExcelMcp.Service/ExcelMcpService.cs)
- [CLI daemon host](../src/ExcelMcp.Service/Rpc/DaemonHost.cs)
- [MCP bridge](../src/ExcelMcp.McpServer/ServiceBridge/ServiceBridge.cs)
- [System map](../CONTEXT.md) and [runtime instructions](agents/rules/architecture-patterns.md)

# ADR-006: Let the MCP SDK own the protocol

**Status:** Current

## Context and decision

The MCP Server uses the official SDK for tool registration, schema handling,
argument binding, transport, and protocol errors. Host filters add
contract-derived argument and session checks. Shared result conversion maps
Service outcomes into SDK tool results.

The server exposes tools. Client-facing consent advice is not a server
confirmation mechanism; the implementation does not make MCP elicitation
requests.

## Reasons and tradeoffs

A custom dispatcher or client-specific schema rewriting would create a second
implementation of protocol behavior to maintain. SDK ownership keeps the host
focused on Excel responsibilities and interoperability.

The host still has to distinguish an operation failure from a valid diagnostic
answer and keep structured results consistent with text results. Protocol
ownership does not transfer Excel execution or session recovery to the SDK.
Using SDK behavior also makes SDK discovery and adapter tests important when
updating tool definitions.

## Implementation and guidance

- [Registration and host filters](../src/ExcelMcp.McpServer/Program.cs)
- [Result conversion](../src/ExcelMcp.McpServer/Tools/ExcelToolsBase.cs)
- [Result contract tests](../tests/ExcelMcp.McpServer.Tests/Integration/Tools/McpResultContractTests.cs)
- [MCP implementation instructions](agents/rules/mcp-server-guide.md)

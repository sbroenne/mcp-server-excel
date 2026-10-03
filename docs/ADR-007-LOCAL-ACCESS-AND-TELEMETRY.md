# ADR-007: Separate workbook access from telemetry

**Status:** Current

## Context and decision

Workbook automation runs under the user's Windows account. CLI connections
use a user-specific named pipe and access checks; MCP communicates with its
client through stdio and invokes its Service in-process.

Requested workbook results go back to the caller. Telemetry is a separate
path carrying restricted operation/outcome information rather than workbook
contents or raw errors. Optional remote formatting and Python in Excel have
separate data flows described in the privacy policy.

## Reasons and tradeoffs

A local Service avoids adding a hosted workbook-processing backend and its
account model. It is not a sandbox: the automation operates with the user's
permissions, and the chosen AI client can receive the requested data.

Raw request and exception logging would make diagnosis easier but could
expose workbook contents, paths, credentials, or queries. Restricted telemetry
provides usage and failure trends at the cost of less remote diagnostic detail.
Local execution does not imply that all features or telemetry are offline.

## Implementation and guidance

- [Pipe access checks](../src/ExcelMcp.Service/ServiceSecurity.cs)
- [CLI telemetry](../src/ExcelMcp.CLI/Telemetry/CliTelemetry.cs)
- [MCP telemetry](../src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs)
- [Privacy policy and data flows](../PRIVACY.md)
- [Security policy](../SECURITY.md)

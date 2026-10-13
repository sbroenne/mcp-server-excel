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

The CLI and MCP Server share a random telemetry user ID stored in the current
Windows user's `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id` file. The file is
created on first use even if telemetry is off. Storage failures use a
process-only random ID without interrupting workbook operations. When the file
is first created with telemetry active, a one-time migration event links the
new ID to the previous name-derived ID so existing usage remains countable.
The weekly analytics queries use that link across retained events but never
publish the legacy ID. Deleting the file resets the stored ID.

## Reasons and tradeoffs

A local Service avoids adding a hosted workbook-processing backend and its
account model. It is not a sandbox: the automation operates with the user's
permissions, and the chosen AI client can receive the requested data.

Raw request and exception logging would make diagnosis easier but could
expose workbook contents, paths, credentials, or queries. Restricted telemetry
provides usage and failure trends at the cost of less remote diagnostic detail.
Local execution does not imply that all features or telemetry are offline.
The random ID remains pseudonymous because it is stable across runs. The
migration hash is sent only once, to the active telemetry destination; if
telemetry is off during first use it is not sent later, and local mode does not
defer it to Azure.

## Implementation and guidance

- [Pipe access checks](../src/ExcelMcp.Service/ServiceSecurity.cs)
- [CLI telemetry](../src/ExcelMcp.CLI/Telemetry/CliTelemetry.cs)
- [MCP telemetry](../src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs)
- [Privacy policy and data flows](../PRIVACY.md)
- [Security policy](../SECURITY.md)

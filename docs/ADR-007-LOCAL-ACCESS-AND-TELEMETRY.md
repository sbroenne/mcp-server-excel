# ADR-007: Separate workbook access from telemetry

**Status:** Current

## Context and decision

Workbook automation runs under the user's Windows account. CLI connections
use a user-specific named pipe and access checks; MCP communicates with its
client through stdio and invokes its Service in-process.

Requested workbook results go back to the caller. Telemetry is a separate
path carrying restricted operation/outcome information rather than workbook
contents or raw errors. MCP unhandled-exception reports keep exception type
names and the owning failure site but redact messages and stacks. Both entry
points correlate events with a stored random ID. It is stable across runs
and remains pseudonymous, not anonymous.
Optional remote formatting and Python in Excel have separate data flows
described in the privacy policy.

The CLI and MCP Server share a random telemetry user ID stored in the current
Windows user's `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id` file. The file is
created/read only by background reporting setup when telemetry is enabled. Storage failures use a
process-only random ID without interrupting workbook operations. When the file
is first created with telemetry active, a one-time migration event links the
new ID to the previous name-derived ID so existing usage remains countable.
The weekly analytics queries use that link across retained events but never
publish the legacy ID. Deleting the file and restarting the MCP Server resets
the stored ID.

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
telemetry is off no identity storage is accessed, and local mode does not
defer it to Azure.

Both entry points use the same native OpenTelemetry batch pipeline, exporting
only explicit usage traces and custom-event logs. Setup and ID storage run in
the background. Product calls skip records before readiness; full buffers drop.
Exit only stops acceptance: it never flushes or disposes reporting providers,
which are not owned by the MCP host. Loss of unsent usage reports is accepted
to keep telemetry off the command, request, startup, and exit critical paths.

CLI reporting belongs to the existing long-lived daemon that performs the
operations, not the short-lived command process. It records each received
operation once and sends queued records while it remains running, without a
new process, extra telemetry request, or client-side flush. Its telemetry
session ID spans multiple CLI invocations, and timings describe service
execution rather than the whole client invocation. Readiness pings and commands
that never reach the daemon are not reported. MCP continues to report in its
own server process; its controller calls do not collect CLI telemetry. Private
Excel workers start before either product's reporting setup and do not create
a second reporting owner.

## Implementation and guidance

- [Pipe access checks](../src/ExcelMcp.Service/ServiceSecurity.cs)
- [CLI telemetry](../src/ExcelMcp.CLI/Telemetry/CliTelemetry.cs)
- [MCP telemetry](../src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs)
- [Privacy policy and data flows](../PRIVACY.md)
- [Security policy](../SECURITY.md)

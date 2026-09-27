# ExcelMcp Architecture

ExcelMcp controls the actual Microsoft Excel desktop application—not just
`.xlsx` files. Windows uses the complete COM backend. macOS x64/Arm64 uses
a capability-gated Apple Events backend for the verified workbook, worksheet,
range, formula, clear, and calculation subset.

## Two equal entry points

The project ships both an MCP Server and a CLI. They are first-class entry
points backed by the same Core commands, parameters, defaults, and validation:

- **MCP Server** hosts `ExcelMcpService` in-process and uses direct method calls,
  which suits conversational and interactive AI clients.
- **CLI** (`excelcli`) communicates with an `ExcelMcpService` background daemon
  over user-local IPC. The daemon keeps workbook sessions open across CLI
  invocations for scripting and coding-agent workflows.

```text
MCP Server ──► In-process ExcelMcpService ──► platform backend ──► Excel
CLI ─────────► CLI daemon (local IPC) ──────► platform backend ──► Excel
```

The entry points run as separate processes, each managing its own Excel
instance. They do not share live sessions.

The Service owns command routing and workbook sessions. A separate Service
daemon host owns named-pipe acceptance, connection limits, idle shutdown, and
connection draining for the CLI process. The MCP Server does not use this pipe:
its bridge owns an in-process Service generation and prevents cancellation from
an older generation from resetting a newer one.

The CLI also avoids loading the MCP tool schemas into a coding agent's context.
In a same-task, same-model benchmark, the CLI workflow used about 59K tokens
versus 163K for MCP—a 64% reduction. Actual usage varies by client, model, and
workflow.

## Core layers

1. **ComInterop** (`src/ExcelMcp.ComInterop`) provides the complete Windows
   backend: STA threading, session management, COM cleanup, write guards, and
   OLE message filtering.
2. **Core** (`src/ExcelMcp.Core`) implements Excel operations for Power Query,
   DAX, VBA, worksheets, ranges, charts, and other domains.
3. **Service** (`src/ExcelMcp.Service`) manages sessions, routes commands, and
   selects the Windows COM or macOS Apple Events backend.
4. **CLI** (`src/ExcelMcp.CLI`) exposes generated command categories and uses a
   persistent daemon.
5. **MCP Server** (`src/ExcelMcp.McpServer`) exposes generated MCP tools and
   invokes the service in-process.
6. **Source generators** (`src/ExcelMcp.Generators*`) generate CLI commands,
   MCP schemas, and skill manifests from Core interfaces.
7. **Optional Office.js bridge** (`office-addin`) provides an action-gated
   localhost HTTPS broker and Excel task pane for future Mac capability tiers.
   It is not required by the base backend and currently exposes health only.

## Platform backends

ExcelMcp intentionally drives Excel rather than rewriting workbook packages.
On Windows, COM provides the complete 326-operation surface, including Power
Query, the Data Model, PivotTables, VBA, charts, and formatting. On macOS,
OSAKit sends bounded Apple Events from the same process identity that performs
the non-prompting Automation permission check. Unsupported actions return
`PlatformNotSupported`.

Windows sessions own an Excel process. macOS sessions own only an exact
workbook inside the user's shared Excel application; they never terminate
Excel or close unrelated workbooks. Existing macOS files are handed to Excel
through LaunchServices and attached by exact path.

The optional Office.js tier is a separate, versioned capability boundary. Its
broker authenticates the local channel, binds an Office.js runtime to the exact
saved workbook URL and ExcelMcp session, serializes requests, enforces
deadlines/cancellation, and negotiates `ExcelApi` requirement sets plus the
running Excel version. No feature is routed through this tier until real Excel
proves the full public contract through both entry points. See
[Office.js bridge setup and security](MACOS-OFFICEJS.md).

## CLI desktop integration

The CLI daemon keeps sessions alive between commands. Windows additionally
provides the system-tray and window-management experience. macOS uses
foreground-safe daemon behavior without Windows UI dependencies.

## Session lifecycle

Both entry points use explicit sessions:

1. Open or create a workbook and receive a session ID.
2. Run one or more operations against that session.
3. Close the session, optionally saving changes.

This avoids repeatedly opening workbooks and gives ExcelMcp one controlled
place to manage platform resources and workbook ownership.

See [macOS support](../specs/MACOS-SUPPORT.md) for the exact capability matrix
and parity evidence.

[Read the development guide](DEVELOPMENT.md) for implementation details, or
[choose an installation path](INSTALLATION.md).

# ExcelMcp Architecture

ExcelMcp controls the actual Microsoft Excel desktop application—not just
`.xlsx` files. Windows uses the complete COM backend. Apple Silicon macOS packages use
a capability-gated Apple Events backend for the verified operation subset in
the [generated action inventory](MACOS-ACTION-INVENTORY.md).
**macOS support is experimental beta**, not full Windows parity.
Desktop acceptance has run on Apple Silicon; Intel Macs are unsupported. See
[macOS support and limitations](../specs/MACOS-SUPPORT.md).

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

The entry points run as separate processes and do not share live sessions.
Windows sessions own separate Excel instances. Mac sessions own exact workbooks
inside shared desktop Excel, not the application or other users' workbooks.

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

The [current decision records](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/DECISIONS.md)
explain the reasons and tradeoffs behind these boundaries.

1. **ComInterop** (`src/ExcelMcp.ComInterop`) provides reusable STA threading,
   session management, COM cleanup, write guards, and OLE message filtering.
2. **Core** (`src/ExcelMcp.Core`) implements Excel operations for Power Query,
   DAX, VBA, worksheets, ranges, charts, and other domains.
3. **Service** (`src/ExcelMcp.Service`) manages sessions, routes commands, and
   selects the Windows COM or macOS Apple Events backend; its daemon host
   provides the CLI's named-pipe process lifetime.
4. **CLI** (`src/ExcelMcp.CLI`) exposes generated command categories and uses a
   persistent daemon.
5. **MCP Server** (`src/ExcelMcp.McpServer`) exposes generated MCP tools and
   invokes the service in-process.
6. **Source generators** (`src/ExcelMcp.Generators*`) generate CLI commands,
   MCP schemas, and skill manifests from Core interfaces.

## Platform backends

Both platforms use the same generated typed dispatch, defaults, and result
serialization. A generated platform command set registers Core interfaces with
Windows commands or Mac transport adapters. Mac batches reject COM callbacks
and references; platform-specific session lifetime remains separate. Remaining
Excel-specific validation and scripting logic are still being migrated.

ExcelMcp intentionally drives Excel rather than rewriting workbook packages.
On Windows, COM provides the complete public operation surface, including Power
Query, the Data Model, PivotTables, VBA, charts, and formatting. On macOS,
session preflight, attachment, open-state checks, close, worksheet listing/creation/rename,
rectangular formula reads/writes, and normal sheet/range calculation use direct C# Apple
Events. Remaining operation routes still use OSAKit while their native
replacements are implemented and verified. Both run in a bounded automation
child with the non-prompting permission check and owning-parent identity
verification. Native calls receive the remaining operation deadline; a native
timeout remains a timeout, not a successful result. Unsupported actions return
`PlatformNotSupported`.

Windows sessions own an Excel process. macOS sessions own only an exact
workbook inside the user's shared Excel application; they never terminate
Excel or close unrelated workbooks. Existing macOS files are handed to Excel
through LaunchServices and attached by exact path.

There is no Office.js add-in or hosted broker. A separately versioned,
user-installed VBA helper is being evaluated for object-model operations that
Excel's Apple Events dictionary cannot expose. Helper installation alone does
not enable an action; public contracts remain gated until desktop parity is
verified.

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

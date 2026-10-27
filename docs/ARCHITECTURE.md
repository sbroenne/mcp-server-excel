# ExcelMcp Architecture

ExcelMcp uses Windows COM automation to control the actual Microsoft Excel
application—not just `.xlsx` files. Because it drives Excel's official
`Excel.Application` API, it can refresh Power Query, recalculate formulas,
refresh PivotTables and the Data Model, evaluate DAX, and run VBA or Python
`=PY()` while preserving existing workbook features.

## Two equal entry points

The project ships both an MCP Server and a CLI. They are first-class entry
points backed by the same Core commands, parameters, defaults, and validation:

- **MCP Server** owns a controller through its Service bridge.
- **CLI** (`excelcli`) communicates with its persistent controller over
  user-isolated Windows named pipes. Sessions remain open across CLI invocations.

Both controllers route Excel work to private workers hosting the same
`ExcelMcpService` and Core commands.

```text
CLI client process
  -> named pipe -> CLI controller -> private worker [Service -> Core] -> Excel COM

MCP Server [Service bridge -> controller] -> private worker [Service -> Core] -> Excel COM
```

### Process lifetimes

| Process | Lifetime | Responsibilities |
| --- | --- | --- |
| CLI client (`excelcli`) | One command invocation, including a batch | Parses arguments, sends requests, prints results, then exits. It does not own workbook sessions or send operation telemetry. |
| CLI daemon (`excelcli service run`) | Persists across client invocations | Owns the CLI controller, tray, recorded status, workers, and operation telemetry. |
| MCP Server | Persists across requests until server shutdown | Owns its controller through the bridge, its workers, and operation telemetry. |
| Private Excel worker | Remains while owning sessions or retained work | Hosts the shared Service, owns Excel COM objects, and serializes work for its Excel application. |

The CLI client and daemon are separate processes running the same executable.
Workers launch the same CLI or MCP executable in a private mode before normal
product startup. `ExcelMcpService` remains a library, not a separately installed
product. CLI and MCP own separate session sets; they do not share live sessions.
Workers start independently of the controller's standard input and output, so a
retained worker cannot keep an MCP client's connection open after controller exit.

The Service owns Excel command routing and workbook sessions inside workers.
The controller records worker status and routes requests without Excel COM.
The shared listener uses .NET hosting, StreamJsonRpc, and protected Windows named
pipes. MCP does not use the CLI controller's pipe; it owns its own controller.

The CLI daemon accepts up to 64 simultaneous connections and up to ten work
requests. It rejects another work request for a session whose request is still
running, rather than letting repeated requests occupy the service. Health,
session-list, and window show/hide requests do not consume the work limit.
Commands that do not use an existing workbook session still consume the work
limit, but do not reserve a workbook session even if batch supplies its ID.
Recovery controls also have a separate pipe so full work connections cannot
exclude them. Incoming data is bounded to 64 MiB per connection before JSON
parsing; normal clients use one request per connection. Each worker accepts up
to eight connections and one Excel operation at a time. Controllers limit the
number of workers to 32. Worker startup, status, release, and connection
draining all have finite waits. These limits do not shorten Excel operation waits.
The daemon remains running when idle until explicitly stopped or its process
exits. An expired caller wait does not stop Excel or undo workbook changes.
An unreachable worker is reported separately from controller health.
The controller continues observing a request after its caller stops waiting,
so a late result still updates the recorded session and workbook path. It does
not send the edit again.
Validation and cross-file worksheet commands reserve each file before
starting a worker, just as opening a workbook does. A rejected request releases
only its own reservations. If both a worker and its recorded Excel process have
exited, the controller releases unusable routes and paths but continues reporting
the old session's last outcome as unknown.
If a helper exits before reporting which Excel process it started, retained
sessions and pending file requests stay protected until explicit recovery or
controller restart. A missing process identity is not proof that Excel closed.
Cross-file worksheet commands register their temporary Excel batch and both
file paths with the existing session manager before waiting for startup.
Success releases that session; failure retains it and reports both files and
the original error. Worker release and controller-loss handling therefore see
unfinished cross-file work as owned work, not an empty helper.
Retained worker records include process start times, controller identity,
generation, and build identity. A new controller requires explicit discovery
and attachment; it never repeats an uncertain edit. A surviving worker reveals
its retained workbooks after controller loss without saving or closing them.
Recovery through another controller requires the same build. After an upgrade,
the retained workbooks remain available for direct inspection in Excel, but
the new build cannot claim their old worker.

Telemetry ownership and delivery tradeoffs are described in
[ADR-007](ADR-007-LOCAL-ACCESS-AND-TELEMETRY.md). A short-lived client's background
task cannot continue sending after that client process exits; the existing
daemon provides the lifetime needed for CLI operation reporting.

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
3. **Service** (`src/ExcelMcp.Service`) manages sessions and routes commands;
   its controller and workers provide the shared execution boundary.
4. **CLI** (`src/ExcelMcp.CLI`) exposes generated command categories and uses a
   persistent daemon.
5. **MCP Server** (`src/ExcelMcp.McpServer`) exposes generated MCP tools and
   calls its own controller, independent of the CLI daemon.
6. **Source generators** (`src/ExcelMcp.Generators*`) generate CLI commands,
   MCP schemas, and skill manifests from Core interfaces.

The generators share parameter aggregation, requiredness, and input-type
projections instead of reconstructing those rules separately. The MCP request
filter checks action applicability and required inputs against that generated
contract, including required inline-value/file alternatives. Inapplicable inputs
set to null or their advertised default are ignored. Worksheet
and screenshot tools retain their specialized handlers but use the same Core
action contract; file lifecycle inputs declare applicability beside their
specialized tool signatures. The SDK still owns schema serialization and
argument binding. Excel-free contract tests compare live SDK discovery with
the runtime input rules, including omitted/null and action-inapplicable inputs.

## Real Excel automation

ExcelMcp intentionally uses the Excel COM API rather than rewriting workbook
packages. This provides:

- Excel's own calculation and refresh engines
- Preservation of formulas, formatting, charts, PivotTables, macros, and the
  Data Model
- Interactive authentication for protected workbooks
- The ability to show Excel and inspect changes as they happen

## CLI desktop integration

The CLI daemon keeps sessions alive between commands and exposes a system-tray
icon for monitoring sessions, update notifications, save prompts, and stopping
the daemon. Excel can remain hidden for speed or be shown and arranged beside an
AI assistant for interactive work.

## Session lifecycle

Both entry points use explicit sessions:

1. Open or create a workbook and receive a session ID.
2. Run one or more operations against that session.
3. Close the session, optionally saving changes.

This avoids repeatedly opening workbooks and gives ExcelMcp one controlled place
to manage COM resources and Excel process shutdown.

[Read the development guide](DEVELOPMENT.md) for implementation details, or
[choose an installation path](INSTALLATION.md).

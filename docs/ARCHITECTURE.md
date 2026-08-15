# ExcelMcp Architecture

ExcelMcp uses Windows COM automation to control the actual Microsoft Excel
application—not just `.xlsx` files. Because it drives Excel's official
`Excel.Application` API, it can refresh Power Query, recalculate formulas,
refresh PivotTables and the Data Model, evaluate DAX, and run VBA or Python
`=PY()` while preserving existing workbook features.

## Two equal entry points

The project ships both an MCP Server and a CLI. They are first-class entry
points backed by the same Core commands, parameters, defaults, and validation:

- **MCP Server** hosts `ExcelMcpService` in-process and uses direct method calls,
  which suits conversational and interactive AI clients.
- **CLI** (`excelcli`) communicates with an `ExcelMcpService` background daemon
  over a user-isolated Windows named pipe. The daemon keeps workbook sessions
  open across CLI invocations for scripting and coding-agent workflows.

```text
MCP Server ──► In-process ExcelMcpService ──► Core Commands ──► Excel COM
CLI ─────────► CLI daemon (named pipe) ─────► Core Commands ──► Excel COM
```

The entry points run as separate processes, each managing its own Excel
instance. They do not share live sessions.

The CLI also avoids loading the MCP tool schemas into a coding agent's context.
In a same-task, same-model benchmark, the CLI workflow used about 59K tokens
versus 163K for MCP—a 64% reduction. Actual usage varies by client, model, and
workflow.

## Core layers

1. **ComInterop** (`src/ExcelMcp.ComInterop`) provides reusable STA threading,
   session management, COM cleanup, write guards, and OLE message filtering.
2. **Core** (`src/ExcelMcp.Core`) implements Excel operations for Power Query,
   DAX, VBA, worksheets, ranges, charts, and other domains.
3. **Service** (`src/ExcelMcp.Service`) manages sessions and routes commands.
4. **CLI** (`src/ExcelMcp.CLI`) exposes generated command categories and uses a
   persistent daemon.
5. **MCP Server** (`src/ExcelMcp.McpServer`) exposes generated MCP tools and
   invokes the service in-process.
6. **Source generators** (`src/ExcelMcp.Generators*`) generate CLI commands,
   MCP schemas, and skill manifests from Core interfaces.

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

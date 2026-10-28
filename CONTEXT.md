# ExcelMcp Context

## Purpose

ExcelMcp is a Windows-only automation system that controls the installed Microsoft Excel application through its COM API. It uses Excel itself to calculate formulas, refresh data, run macros, and preserve workbook features that file-only tools cannot safely reproduce.

## System map

```text
CLI client process (one invocation)
  -> named pipe -> CLI controller -> private worker [ExcelMcpService -> Core] -> Excel COM

MCP Server [owned Service bridge -> controller]
  -> private worker [ExcelMcpService -> Core] -> Excel COM
```

The MCP Server and `excelcli` are equal user entry points. They expose the same
operations and behavior, but the CLI client exits after its invocation while
the daemon persists across invocations. Each host owns its own controller,
separate workbook sessions, and operation telemetry. Excel execution lives in
private workers launched from that entry point, not in the controller or tray.
`ExcelMcpService` is a library hosted inside each worker. See
[process lifetimes](docs/ARCHITECTURE.md#process-lifetimes).

## Glossary

- **Entry point:** The MCP Server or `excelcli`, through which a user or agent requests Excel work.
- **Session:** A managed connection to a workbook and its Excel process. It either owns an Excel process it opened or attaches to an already-open workbook without reopening it. A session stays open across operations until it is closed.
- **Session ID:** The identifier returned when a workbook is opened, created, or attached. Later operations use it to select the session.
- **Batch (`IExcelBatch`):** The internal object that keeps Excel and its workbook open and runs COM work on Excel's required thread.
- **Core command:** Transport-independent Excel behavior implemented under `src/ExcelMcp.Core`.
- **Service:** The shared command router and session owner used by both entry points.
- **Controller:** The host-owned router that tracks workers, workbook sessions,
  and process identities. Basic status reads recorded state without waiting for Excel.
- **Worker:** A private copy of the current product process that hosts the Service
  and owns one Excel application's execution. It is not a third product entry point.
- **Daemon host:** The shared named-pipe listener, with bounded connections and
  shutdown draining. The CLI controller stays running until explicitly stopped.
- **Service bridge:** The MCP host component that owns one controller
  generation and prevents stale requests from disposing a newer generation.
- **COM reference:** A live Excel object such as a workbook, worksheet, range, chart, or model object. It belongs to the Excel process and requires controlled cleanup.
- **Generated surface:** CLI commands, service routes, MCP schemas, or reference material produced from a source contract rather than maintained separately.
- **Source contract:** An annotated Core interface from which matching Service, CLI, and MCP behavior is generated.
- **Worksheet table:** An Excel table visible on a worksheet.
- **Data Model table:** A table loaded into Excel's internal analytical model. It is separate from its worksheet source.
- **Regular PivotTable:** A PivotTable backed by worksheet data or a normal PivotCache.
- **OLAP/Data Model PivotTable:** A PivotTable backed by Excel's Data Model and addressed through OLAP fields and measures.
- **Linked PivotChart:** A chart whose `PivotLayout` points to its source PivotTable and continues to follow PivotTable changes.
- **Power Query:** Excel's query and transformation engine. A query may load to a worksheet, the Data Model, both, or remain connection-only.

## Runtime relationships

- One session operates on one workbook in an Excel application process. An owned session starts that process; an attached session leaves the user's application and neighboring workbooks open when detached.
- Operations inside one session run in order on one Excel thread.
- Different Excel applications can run independently. Workbooks sharing an
  application share one execution owner; the same workbook cannot be opened
  in multiple sessions.
- An expired wait can leave Excel working. Retain the session and workbook, reveal the affected Excel window when possible, and report status to the calling agent. New work is refused until the original callback has returned and Excel readiness is confirmed; then the same session can be reused.
- Ordinary operations change the in-memory workbook; Excel AutoSave can persist edits independently. Opening and attaching preserve that setting. Explicit close defaults to discarding unsaved changes; request saving to keep them. Normal Service shutdown closes confirmed-saved workbooks it opened and hands unsaved ones back open and visible without saving or discarding. Attached workbooks stay untouched. An expired wait or caller cancellation does not authorize saving, discarding, rollback, or termination.
- MCP uses the official SDK for registration, transport, argument binding, and protocol errors. The host owns its injected Service bridge; cancelled startup retains its eventual session. Session publication is coordinated with shutdown.

## Sources of truth

- Root and nested `AGENTS.md` files define contributor rules beside their owners. The root task map requires cross-area guides explicitly; clients do not all discover descendant files automatically. The root [Code Review Rules](AGENTS.md#code-review-rules) section covers review tasks; [runtime boundaries](src/AGENTS.md), [COM safety](src/ExcelMcp.ComInterop/AGENTS.md), and [MCP boundaries](src/ExcelMcp.McpServer/AGENTS.md) cover implementation.
- [Architecture decisions](docs/DECISIONS.md) explain the reasons and tradeoffs behind current choices, without duplicating those instructions.
- `docs/ARCHITECTURE.md` explains the public architecture.
- Annotated Core interfaces and their implementations define operation contracts and behavior.
- `docs/features/` documents user-facing behavior.
- `docs/reference/` owns general workflows, limitations, and recovery documentation.
- `skills/` contains a small CLI launcher-discovery skill and the CLI and MCP
  report-formatting skills; packaging selects the formatting skills' reference
  from `docs/reference/report-formatting.md`.

When these sources disagree, confirm the current implementation and update the stale source instead of creating another competing definition.

Track proposed feature requirements in GitHub issues. Do not maintain separate
specification copies of contracts or shipped feature documentation.

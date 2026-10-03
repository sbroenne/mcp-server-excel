# ExcelMcp Context

## Purpose

ExcelMcp is a Windows-only automation system that controls the installed Microsoft Excel application through its COM API. It uses Excel itself to calculate formulas, refresh data, run macros, and preserve workbook features that file-only tools cannot safely reproduce.

## System map

```text
MCP Server -> owned Service bridge -> in-process ExcelMcpService -> Core -> Excel COM
CLI parser -> named-pipe daemon host -> ExcelMcpService -> Core -> Excel COM
```

The MCP Server and `excelcli` are equal user entry points. They expose the same operations and behavior, but they run in separate processes and do not share open sessions.

## Glossary

- **Entry point:** The MCP Server or `excelcli`, through which a user or agent requests Excel work.
- **Session:** A managed connection to an open workbook and its Excel process. A session stays open across operations until it is closed.
- **Session ID:** The identifier returned when a workbook is opened or created. Later operations use it to select the session.
- **Batch (`IExcelBatch`):** The internal object that keeps Excel and its workbook open and runs COM work on Excel's required thread.
- **Core command:** Transport-independent Excel behavior implemented under `src/ExcelMcp.Core`.
- **Service:** The shared command router and session owner used by both entry points.
- **Daemon host:** The CLI process component that owns named-pipe acceptance,
  connection limits, idle shutdown, and connection draining around the Service.
- **Service bridge:** The MCP host component that owns one in-process Service
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

- One session owns one Excel application process and may contain one or more open workbooks.
- Operations inside one session run in order on one Excel thread.
- Different sessions can run independently, but the same workbook cannot be opened in multiple sessions.
- A timeout can leave Excel busy after the caller stops waiting. Such a session is no longer safe for additional work and must be closed.
- Ordinary operations change the in-memory workbook. Explicit close defaults to discarding unsaved changes; request saving to keep them. Normal Service shutdown attempts to save remaining sessions before disposal. Bare batch disposal, crashes, and forced cancellation cleanup do not guarantee saving.
- MCP uses the official SDK for registration, transport, argument binding, and protocol errors. The host owns its injected Service bridge; cancelled startup reclaims only the eventual session. Session publication is coordinated with shutdown.

## Sources of truth

- `AGENTS.md`, nested `AGENTS.md` files, and `docs/agents/rules/` define shared coding, testing, COM safety, and release rules. The root [Code Review Rules](AGENTS.md#code-review-rules) section covers review tasks.
- [Architecture decisions](docs/DECISIONS.md) explain the reasons and tradeoffs behind current choices, without duplicating those instructions.
- `docs/ARCHITECTURE.md` explains the public architecture.
- Annotated Core interfaces and their implementations define operation contracts and behavior.
- `docs/features/` documents user-facing behavior.
- `docs/reference/` owns general workflows, limitations, and recovery documentation.
- `skills/` contains only the CLI and MCP report-formatting skills; packaging
  selects their formatting reference from `docs/reference/report-formatting.md`.

When these sources disagree, confirm the current implementation and update the stale source instead of creating another competing definition.

Track proposed feature requirements in GitHub issues. Do not maintain separate
specification copies of contracts or shipped feature documentation.

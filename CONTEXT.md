# ExcelMcp Context

## Purpose

ExcelMcp is a cross-platform desktop Excel automation system. Windows uses the
complete COM backend. Experimental beta Apple Silicon macOS support uses a capability-gated Apple Events
backend for the verified operation subset. Intel macOS is unsupported. Both use Excel itself rather
than a file-only calculation engine.

## System map

```text
MCP Server -> owned Service bridge -> in-process ExcelMcpService -> platform backend -> desktop Excel
CLI parser -> named-pipe daemon host -> ExcelMcpService -> platform backend -> desktop Excel
```

The MCP Server and `excelcli` are equal user entry points. They expose the same operations and behavior, but they run in separate processes and do not share open sessions.

For every action enabled on macOS, the Mac and Windows backends implement the
same public contract 1:1: inputs, defaults, validation, results, workbook
effects, persistence, and errors. Only the supported Excel automation target
and platform ownership mechanics differ. An action or explicitly documented
variant stays gated when the Mac APIs cannot preserve that exact contract.

## Glossary

- **Entry point:** The MCP Server or `excelcli`, through which a user or agent requests Excel work.
- **Session:** A managed connection to an open workbook. On Windows it owns an
  Excel process; on macOS it owns only the exact workbook inside shared desktop
  Excel. A session stays open across operations until it is closed.
- **Session ID:** The identifier returned when a workbook is opened or created. Later operations use it to select the session.
- **Batch (`IExcelBatch`):** The command execution context for an owned workbook.
  Windows runs COM callbacks on Excel's required thread. The internal Mac adapter
  sends bounded automation requests and explicitly rejects COM callbacks and references.
- **Core command:** A Windows Excel behavior implementation under
  `src/ExcelMcp.Core`; annotated interfaces also define the shared generated
  public contract used by both platforms.
- **Mac backend:** The Service adapter that sends bounded Apple Events and
  explicitly gates operations that do not have verified parity. Session
  lifecycle, worksheet listing/creation/rename, rectangular formula reads/writes, and
  normal sheet/range calculation use direct C# native events; remaining commands use OSAKit while
  their native replacements are being implemented and verified.
- **Service:** The shared generated command router and platform session owner used by both entry points.
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

- On Windows, one session owns one Excel application process and may contain one
  or more open workbooks.
- On macOS, one session owns one exact workbook in shared desktop Excel and must
  never terminate Excel or close unknown user workbooks.
- Operations inside one session run in order on one Excel thread.
- Both platforms use the same generated typed argument binding, defaults, and
  result serialization. Platform command sets select the Windows implementation
  or Mac transport adapter behind each Core interface. Excel-specific validation
  and remaining scripting logic still need native/shared-semantic migration.
- Different sessions have separate public namespaces. The same workbook cannot
  be opened in multiple sessions; macOS also serializes open preflight,
  LaunchServices handoff, and attachment across participating processes.
- A timeout can leave Excel busy after the caller stops waiting. Such a session
  is no longer safe for additional work. On macOS, an unconfirmed file-open
  handoff requires manual reconciliation, not automatic close or rollback;
  see `specs/MACOS-SUPPORT.md`.
- Ordinary operations change the in-memory workbook. Explicit close defaults to discarding unsaved changes; request saving to keep them. Normal Service shutdown attempts to save remaining sessions before disposal. Bare batch disposal, crashes, and forced cancellation cleanup do not guarantee saving.
- MCP uses the official SDK for registration, transport, argument binding, and protocol errors. The host owns its injected Service bridge; cancelled startup reclaims only the eventual session. Session publication is coordinated with shutdown.

## Sources of truth

- `AGENTS.md`, nested `AGENTS.md` files, and `docs/agents/rules/` define shared coding, testing, COM safety, and release rules. The root [Code Review Rules](AGENTS.md#code-review-rules) section covers review tasks.
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

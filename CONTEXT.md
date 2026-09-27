# ExcelMcp Context

## Purpose

ExcelMcp is a cross-platform desktop Excel automation system. Windows uses the
complete COM backend. Apple Silicon macOS uses a capability-gated Apple Events
backend for the verified initial operation subset. Both use Excel itself rather
than a file-only calculation engine.

## System map

```text
MCP Server -> in-process ExcelMcpService -> platform backend -> desktop Excel
CLI        -> background ExcelMcpService -> platform backend -> desktop Excel
```

The MCP Server and `excelcli` are equal user entry points. They expose the same operations and behavior, but they run in separate processes and do not share open sessions.

## Glossary

- **Entry point:** The MCP Server or `excelcli`, through which a user or agent requests Excel work.
- **Session:** A managed connection to an open workbook. On Windows it owns an
  Excel process; on macOS it owns only the exact workbook inside shared desktop
  Excel. A session stays open across operations until it is closed.
- **Session ID:** The identifier returned when a workbook is opened or created. Later operations use it to select the session.
- **Batch (`IExcelBatch`):** The Windows internal object that keeps Excel and its
  workbook open and runs COM work on Excel's required thread.
- **Core command:** A Windows Excel behavior implementation under
  `src/ExcelMcp.Core`; annotated interfaces also define the shared generated
  public contract used by both platforms.
- **Mac backend:** The Service adapter that sends bounded Apple Events through
  OSAKit and explicitly gates operations that do not have verified parity.
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

- On Windows, one session owns one Excel application process and may contain one
  or more open workbooks.
- On macOS, one session owns one exact workbook in shared desktop Excel and must
  never terminate Excel or close unknown user workbooks.
- Operations inside one session run in order on one Excel thread.
- Different sessions have separate public namespaces. The same workbook cannot
  be opened in multiple sessions; macOS also serializes open preflight,
  LaunchServices handoff, and attachment across participating processes.
- A timeout can leave Excel busy after the caller stops waiting. Such a session is no longer safe for additional work and must be closed.
- Workbook changes are not automatically saved when a batch or session is disposed. Saving is an explicit operation.

## Sources of truth

- `.github/copilot-instructions.md` and `.github/instructions/` define coding, testing, COM safety, and release rules.
- `docs/ARCHITECTURE.md` explains the public architecture.
- `specs/` defines feature contracts and intended behavior.
- `docs/features/` documents user-facing behavior.
- `skills/shared/` is the source for guidance shared by the generated CLI and MCP skills.

When these sources disagree, confirm the current implementation and update the stale source instead of creating another competing definition.

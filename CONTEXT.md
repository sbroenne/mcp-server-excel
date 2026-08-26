# ExcelMcp Context

## Purpose

ExcelMcp is a Windows-only automation system that controls the installed Microsoft Excel application through its COM API. It uses Excel itself to calculate formulas, refresh data, run macros, and preserve workbook features that file-only tools cannot safely reproduce.

## System map

```text
MCP Server -> in-process ExcelMcpService -> Core commands -> Excel COM
CLI        -> background ExcelMcpService -> Core commands -> Excel COM
```

The MCP Server and `excelcli` are equal user entry points. They expose the same operations and behavior, but they run in separate processes and do not share open sessions.

## Glossary

- **Entry point:** The MCP Server or `excelcli`, through which a user or agent requests Excel work.
- **Session:** A managed connection to an open workbook and its Excel process. A session stays open across operations until it is closed.
- **Session ID:** The identifier returned when a workbook is opened or created. Later operations use it to select the session.
- **Batch (`IExcelBatch`):** The internal object that keeps Excel and its workbook open and runs COM work on Excel's required thread.
- **Core command:** Transport-independent Excel behavior implemented under `src/ExcelMcp.Core`.
- **Service:** The shared command router and session owner used by both entry points.
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
- Workbook changes are not automatically saved when a batch or session is disposed. Saving is an explicit operation.

## Sources of truth

- `.github/copilot-instructions.md` and `.github/instructions/` define coding, testing, COM safety, and release rules.
- `docs/ARCHITECTURE.md` explains the public architecture.
- `specs/` defines feature contracts and intended behavior.
- `docs/features/` documents user-facing behavior.
- `skills/shared/` is the source for guidance shared by the generated CLI and MCP skills.

When these sources disagree, confirm the current implementation and update the stale source instead of creating another competing definition.

# Excel COM API Coverage Specification

> Roadmap for expanding ExcelMcp from high-value automation coverage toward comprehensive,
> deterministic Excel COM coverage.

## Objective

Expose every practical Excel desktop automation capability that:

1. Is available through the supported Excel PIA or a documented late-bound COM surface.
2. Can run deterministically in an unattended MCP Server or CLI workflow.
3. Can be verified with a real Excel integration test.
4. Does not require weakening Office security settings or accepting interactive UI prompts.

This is not a promise to wrap every historical member in the Excel type library. UI-only,
deprecated, cloud-service, callback/event, and security-sensitive APIs are inventoried and
explicitly excluded when they cannot satisfy the criteria above.

## Definition of Complete

An implemented operation is complete only when all of the following are true:

- The Core interface and implementation expose the behavior.
- MCP Server and CLI surfaces are generated from the same Core contract.
- Parameter names, defaults, validation, and result shapes match across both entry points.
- A real Excel Core integration test verifies the resulting workbook state.
- A real MCP integration test verifies transport, schema, dispatch, and result serialization.
- Acquired COM objects are released in `finally` blocks.
- The implementation uses typed PIA APIs unless a compile probe proves a PIA gap.
- User-facing references, skills, operation counts, and a changeset are updated.
- Unsupported members are recorded with a concrete reason instead of a placeholder.

## Coverage Workstreams

| Workstream | Scope | Dependency | Completion evidence |
|---|---|---|---|
| Workbook lifecycle | Built-in/custom properties, Save As, fixed-format export, printing, external links | File/session architecture | Round-trip files, properties, links, and exported output |
| Worksheet views | Freeze/split panes, zoom, headings/gridlines, grouping/outlines, hyperlinks | Workbook window selection | Read-after-write window and worksheet state |
| Drawing objects | Images, shapes, text boxes, connectors, controls, sparklines | Existing worksheet style commands | Lifecycle and formatting round trips |
| What-if analysis | Goal Seek, scenarios, supported analysis APIs, Solver capability detection | Range/formula commands | Calculated cell and scenario state |
| PivotTables and charts | Cache/source management, grouping, drill-through, missing chart types and formatting | Existing PivotTable/chart strategies | Real PivotTable/chart state |
| Data Model | Remaining model metadata, connections, calculated-column feasibility | Power Query/Data Model setup | Real model state or compile-proven exclusion |
| Collaboration and imports | Threaded comments, QueryTables, text/web imports, refresh status/cancellation | Connection and worksheet commands | Real imported data and refresh state |
| Legacy/admin COM | XML maps and deterministic legacy workbook APIs | File and worksheet commands | XML round trip or documented exclusion |

## Delivery Order

1. Complete the independent workstreams in parallel.
2. Integrate shared contracts and result models, resolving action-name and schema conflicts.
3. Regenerate MCP, CLI, skills, and operation counts from the combined Core surface.
4. Run targeted feature tests for each workstream.
5. Run the Excel-dependent CLI and MCP end-to-end smoke workflow.
6. Run the repository's mandatory pre-commit gates before any commit.

## Explicit Exclusion Categories

These categories are inventoried but are not wrapped merely to claim numerical coverage:

| Category | Exclusion rule |
|---|---|
| Application/workbook events | Exclude callback registration unless lifecycle and teardown can be made deterministic across MCP/CLI calls |
| Interactive dialogs and print preview | Exclude operations that block waiting for user input; provide non-interactive equivalents where available |
| Command bars and Ribbon customization | Exclude deprecated or host-UI customization unrelated to workbook automation |
| Mail/envelope APIs | Exclude APIs that depend on Outlook profiles, user credentials, or interactive security prompts |
| Coauthoring/cloud collaboration | Exclude service-backed behavior not exposed through local Excel COM |
| ActiveX and executable controls | Exclude creation or code injection that expands the workbook's executable attack surface |
| Security configuration | Never enable VBA trust, lower macro security, bypass Protected View, or alter Trust Center settings |
| Version-specific members | Implement only with capability detection and a stable failure result when unavailable |
| Unavailable PIA members | Use justified late binding only when a compile probe proves the runtime COM member exists |

## Test Requirements

Every new action requires:

1. A test written before implementation that fails for the missing behavior.
2. A Core integration test opening a real workbook through `ExcelSession.BeginBatch`.
3. A state assertion against Excel, not only `Success`.
4. An MCP integration test invoking the generated tool action.
5. A CLI parity check through generated command coverage.
6. Feature-specific execution with an explicit terminal timeout.

Tests must use unique workbooks and must not save unless persistence is the behavior under
test. APIs requiring unavailable add-ins, credentials, cloud services, or interactive UI are
validated as capability checks and documented exclusions rather than conditionally passing tests.

## Integration Risks

- Several workstreams add result types to the shared result-model file.
- Worksheet views and workbook views both operate on an Excel `Window`; implementations must
  select the window belonging to the session workbook rather than relying on unrelated active state.
- QueryTables and connection refresh operations share COM objects and cancellation semantics.
- PivotTable/chart work must extend existing strategies instead of adding parallel implementations.
- Generated MCP/CLI surfaces make action names public contracts; duplicate or ambiguous actions
  must be resolved before integration.
- Documentation counts are updated only after the combined generated surface is counted.

## Final Verification

The expansion is complete when every workstream is either:

- implemented with Core and MCP end-to-end tests passing, or
- listed as excluded with the exact COM/PIA, security, interactivity, dependency, or reliability
  reason that prevents a production-quality implementation.

The final coverage report must distinguish implemented operations, partial/version-dependent
operations, and exclusions. It must not describe ExcelMcp as a literal one-to-one wrapper unless
all non-excluded members have been verified.

# ADR-002: Use desktop Excel as the workbook engine

**Status:** Current

## Context and decision

ExcelMcp automates the installed Windows Excel application through COM rather
than implementing workbook operations by reading or rewriting file packages.
Typed Excel interop definitions are embedded at build time and are the
intended way to reach Excel. Late binding is accepted only where the typed
definitions have a proven gap or the API belongs to an object model that is
not referenced, such as the VBA editor, ADO, and Office-core enums.

The code has not fully reached that state: several command areas still hold
objects the typed definitions support as `dynamic`. That is conversion debt
tracked in [#1140](https://github.com/sbroenne/mcp-server-excel/issues/1140),
not an accepted exception.

The pre-open protection check reads binary container metadata, not workbook
contents. That exception supports opening protected files; it is not another
workbook engine.

## Reasons and tradeoffs

File libraries can work without desktop Excel and are easier to host across
platforms. They do not provide Excel's full calculation, refresh, Data Model,
macro, and authentication behavior. Using Excel also avoids reconstructing
workbook features the automation did not intend to change.

Typed interop exposes compile-time mistakes that broad dynamic access would
defer until runtime, where they surface as binder exceptions. Limited late
binding accommodates real gaps without adding deployment dependencies for
every Office API.

The cost is a Windows/Excel dependency, COM lifetime management, and features
whose availability depends on the installed Excel version and desktop.

## Implementation and guidance

- [Excel batch](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Protection metadata reader](../src/ExcelMcp.ComInterop/OleCompoundFileReader.cs)
- [Interop project](../src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj)
- [COM instructions](../src/ExcelMcp.ComInterop/AGENTS.md) and [PIA coverage](PIA-COVERAGE.md)

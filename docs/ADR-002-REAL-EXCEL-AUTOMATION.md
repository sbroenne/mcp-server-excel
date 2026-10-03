# ADR-002: Use desktop Excel as the workbook engine

**Status:** Current

## Context and decision

ExcelMcp automates the installed Windows Excel application through COM rather
than implementing workbook operations by reading or rewriting file packages.
Typed Excel interop definitions are embedded at build time; runtime-only gaps
use narrowly scoped late binding instead of requiring unavailable Office
assemblies.

The pre-open protection check reads binary container metadata, not workbook
contents. That exception supports opening protected files; it is not another
workbook engine.

## Reasons and tradeoffs

File libraries can work without desktop Excel and are easier to host across
platforms. They do not provide Excel's full calculation, refresh, Data Model,
macro, and authentication behavior. Using Excel also avoids reconstructing
workbook features the automation did not intend to change.

Typed interop exposes compile-time mistakes that broad dynamic access would
defer until runtime. Limited late binding accommodates gaps without adding
deployment dependencies for every Office API.

The cost is a Windows/Excel dependency, COM lifetime management, and features
whose availability depends on the installed Excel version and desktop.

## Implementation and guidance

- [Excel batch](../src/ExcelMcp.ComInterop/Session/ExcelBatch.cs)
- [Protection metadata reader](../src/ExcelMcp.ComInterop/OleCompoundFileReader.cs)
- [Interop project](../src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj)
- [COM instructions](agents/rules/excel-com-interop.md) and [PIA coverage](PIA-COVERAGE.md)

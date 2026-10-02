# ExcelMcp Office.js bridge

**Experimental beta infrastructure:** installing this package does not enable
the currently unsupported Mac feature actions. See
[macOS beta limitations](../specs/MACOS-SUPPORT.md#not-supported-in-the-macos-beta).

This package is an optional, versioned foundation for macOS Excel capability
tiers. It includes candidate `Excel.run` handlers for tables and table
columns, conditional formatting, same-workbook worksheet copy/move, regular
charts and chart configuration, ordinary local PivotTable creation/deletion,
and source-identifiable slicer creation. The installed configuration enables
only `bridge.health` until live Excel parity evidence exists. The base Apple
Events backend remains usable when the add-in is absent.

Linked PivotCharts, OLAP/Data Model pivots, cache configuration, grouping,
calculated fields/members, and slicer operations whose source cannot be
verified remain outside this tier rather than being approximated.

Disabled internal `ExcelApiDesktop 1.1` handlers can prepare exact range and
window screen geometry and restore the prior workbook view with an opaque
single-use token. They do not identify the macOS process or ScreenCaptureKit
window and remain unavailable until native identity and coordinate mapping are
proved in real Excel. Geometry preparation is fail-closed for ranges that do
not fit in one contained window rectangle; it does not approximate the Windows
tiling planner.

The bridge binds only to `127.0.0.1`, requires HTTPS, authenticates every API
request, validates task-pane origins, binds sessions to exact workbook URLs,
serializes requests per workbook, and expires requests at their deadline.
Numbered `ExcelApi` and `ExcelApiDesktop` support are negotiated independently
at runtime; `ExcelApiDesktop` is not an XML manifest activation requirement.

See [macOS Office.js setup](../docs/MACOS-OFFICEJS.md) for installation,
activation, health, upgrade, and removal commands.

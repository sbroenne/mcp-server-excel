# Change Log

All notable changes to the ExcelMcp VS Code extension will be documented in this file.

## [Unreleased]

- **OLAP PivotChart Creation** (#267): `CreateFromPivotTable` now works with OLAP/Data Model PivotTables

## [1.4.31] - 2025-12-04

- **Locale-Independent Number Formatting** (#263): Number and date formats now work correctly on non-US locales (German, French, etc.)

## [1.4.30] - 2025-12-04

- **OLAP PivotTable AddValueField** (#261): Fixed errors when adding value fields to Data Model PivotTables
- **Show Excel Mode**: Open with `showExcel: true` to watch AI changes live in real-time

## [1.4.28] - 2025-12-01

- **VS Code Extension Display Name** (#257): Corrected MCP server display name

## [1.4.25] - 2025-12-01

- **89% Smaller Extension Size** (#250): Switched to framework-dependent deployment (requires .NET 8.0 runtime)

## [1.4.24] - 2025-12-01

- **Session Stability** (#245): Fixed Excel MCP Server stopping due to network errors
- **PivotTable Grand Totals Control**: Show/hide row and column grand totals independently
- **PivotTable Grouping**: Group dates by days/months/quarters/years, or create numeric ranges
- **PivotTable Calculated Fields**: Create calculated fields with formulas
- **PivotTable Layout & Subtotals**: Configure layout form and subtotals visibility
- Total operations increased to 172

## [1.0.0] - 2025-10-29

- Initial release of ExcelMcp VS Code extension


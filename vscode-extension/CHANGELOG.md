# Change Log

All notable changes to the ExcelMcp VS Code extension will be documented in this file.

## [Unreleased]

### Added
- **PivotTable Grand Totals Control**: New SetGrandTotals operation for independent row and column grand totals control
  - Show/hide row grand totals (bottom summary row)
  - Show/hide column grand totals (right summary column)
  - Independent control - configure row and column separately
  - Full support for both regular and OLAP PivotTables
- **PivotTable Grouping**: Date and numeric grouping operations
  - GroupByDate: Group dates by days, months, quarters, years
  - GroupByNumeric: Create custom numeric ranges with configurable start/end/interval
- **PivotTable Calculated Fields**: Create calculated fields with DAX-like formulas
- **PivotTable Layout & Subtotals**: Configure layout form and subtotals visibility

### Changed
- Total operations increased to 164 (was 163)
- PivotTable tool now has 21 actions (was 20)

## [1.0.0] - 2025-10-29

### Optimization
- Refactored tools to reduce cognizant load on LLM
- Fix stability issues when working on multiple files at the same time

## [1.0.0] - 2025-10-29

### Added
- Improved LLM instructions and optimized PowerQuery Tool

## [1.0.0] - 2025-10-29

### Added
- Initial release of ExcelMcp VS Code extension


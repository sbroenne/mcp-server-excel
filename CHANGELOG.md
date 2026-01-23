# Changelog

All notable changes to ExcelMcp will be documented in this file.

This changelog covers all components:
- **MCP Server** - Model Context Protocol server for AI assistants
- **CLI** - Command-line interface for scripting
- **VS Code Extension** - One-click installation with bundled MCP Server
- **MCPB** - Claude Desktop bundle for one-click installation

## [Unreleased]

### Added
- **Chart Formatting** (#384): Enhanced chart formatting capabilities
  - **Data Labels**: Configure label position and visibility (showValue, showCategory, showPercentage, etc.)
  - **Axis Scale**: Get/set axis scale properties (min, max, units, auto-scale flags)
  - **Gridlines**: Control major/minor gridlines visibility on chart axes
  - **Series Markers**: Configure marker style, size, and colors for data series
  - 8 new operations bringing total chart operations to 22

## [1.5.11] - 2025-01-22

### Added
- Added Agent Skill to all artifacts

## [1.5.9] - 2025-01-20

### Fixed
- **CreateEmpty Error Handling** (#372): File creation errors now return proper JSON with `isError: true` instead of crashing

### Changed
- **MCPB Submission Compliance**: Bundle now includes LICENSE and CHANGELOG.md per Anthropic requirements
- **Documentation Updates**: All READMEs updated with LLM-tested example prompts and accurate tool counts (22 tools, 194 operations)

## [1.5.8] - 2025-01-20

### Added
- Now available as a Claude Desktop MCPB Extension
  
## [1.5.6] - 2025-01-20

### Added
- **PivotTable & Table Slicers** (#363): New `excel_slicer` tool for interactive filtering
  - **PivotTable Slicers**: Create, list, filter, and delete slicers for PivotTable fields
  - **Table Slicers**: Create, list, filter, and delete slicers for Excel Table columns
  - 8 new operations for interactive data filtering

## [1.5.5] - 2025-01-19

### Added
- **DMV Query Execution** (#353): Query Data Model metadata using Dynamic Management Views
  - New `execute-dmv` action on `excel_datamodel` tool
  - Query TMSCHEMA_MEASURES, TMSCHEMA_RELATIONSHIPS, DISCOVER_CALC_DEPENDENCY, etc.

## [1.5.4] - 2025-01-19

### Added
- **DAX EVALUATE Query Execution** (#356): Execute DAX queries against the Data Model
  - New `evaluate` action on `excel_datamodel` tool for ad-hoc DAX queries
- **DAX-Backed Excel Tables** (#356): Create worksheet tables populated by DAX queries
  - New `create-from-dax`, `update-dax`, `get-dax` actions

## [1.5.0] - 2025-01-10

### Changed
- **Tool Reorganization** (#341): Split 12 monolithic tools into 21 focused tools
  - 186 operations total, better organized for AI assistants
  - Ranges: 4 tools (excel_range, excel_range_edit, excel_range_format, excel_range_link)
  - PivotTables: 3 tools (excel_pivottable, excel_pivottable_field, excel_pivottable_calc)
  - Tables: 2 tools (excel_table, excel_table_column)
  - Data Model: 2 tools (excel_datamodel, excel_datamodel_rel)
  - Charts: 2 tools (excel_chart, excel_chart_config)
  - Worksheets: 2 tools (excel_worksheet, excel_worksheet_style)

### Added
- **LLM Integration Testing** (#341): Real AI agent testing using [agent-benchmark](https://github.com/mykhaliev/agent-benchmark)

### Changed
- **.NET 10 Upgrade**: Requires .NET 10.0 instead of .NET 8.0

## [1.4.42] - 2025-12-15

### Added
- **Power Query Rename** (#326, #327): New `rename` action for Power Query queries
- **Data Model Table Rename** (#326, #327): New `rename-table` action for Data Model tables

## [1.4.41] - 2025-12-14

### Fixed
- **Power Query Data Model Fix** (#324): Fixed "0x800A03EC" error when updating Power Query in workbooks with Data Model present

## [1.4.40] - 2025-12-14

### Changed
- **MCP SDK Upgrade** (#301): Upgraded ModelContextProtocol SDK from 0.4.1-preview.1 to 0.5.0-preview.1
  - Proper `isError` signaling for tool execution failures
  - Deterministic exit codes (0 = success, 1 = fatal error)

## [1.4.37] - 2025-12-06

### Changed
- **PivotTable Performance** (#286): Optimized `RefreshTable()` calls

### Added
- **Data Model Members** (#288): Added support for Data Model table members

## [1.4.36] - 2025-12-06

### Changed
- **Documentation Updates** (#290): Updated tool/operation counts

### Fixed
- **SEO Fix** (#292): Fixed robots.txt sitemap URL

## [1.4.35] - 2025-12-05

### Added
- **Data Model Relationships** (#278): Full support for creating, updating, and deleting relationships
- **Custom Domain** (#276): excelmcpserver.dev

## [1.4.34] - 2025-12-05

### Fixed
- **DAX Formula Locale Handling** (#281): DAX formulas now work on European locales

## [1.4.33] - 2025-12-04

### Changed
- **Atomic Cross-File Worksheet Operations** (#273): New `copy-to-file` and `move-to-file` actions

## [1.4.32] - 2025-12-04

### Fixed
- **OLAP PivotChart Creation** (#267): `CreateFromPivotTable` now works with OLAP/Data Model PivotTables
- **Power Query LoadToBoth Detection** (#271): Fixed incorrect detection

## [1.4.31] - 2025-12-04

### Fixed
- **Locale-Independent Number Formatting** (#263): Number and date formats now work on non-US locales

## [1.4.30] - 2025-12-03

### Fixed
- **OLAP PivotTable AddValueField** (#261): Fixed errors when adding value fields to Data Model PivotTables

### Added
- **Show Excel Mode**: Open with `showExcel: true` to watch AI changes live

## [1.4.28] - 2025-12-01

### Fixed
- **VS Code Extension Display Name** (#257): Corrected MCP server display name

## [1.4.25] - 2025-12-01

### Changed
- **89% Smaller Extension Size** (#250): Switched to framework-dependent deployment

## [1.4.24] - 2025-12-01

### Fixed
- **Session Stability** (#245): Fixed Excel MCP Server stopping due to network errors

### Added
- **PivotTable Grand Totals Control**: Show/hide row and column grand totals
- **PivotTable Grouping**: Group dates by days/months/quarters/years
- **PivotTable Calculated Fields**: Create calculated fields with formulas
- **PivotTable Layout & Subtotals**: Configure layout form and subtotals visibility
- Total operations: 172

## [1.4.0] - 2025-11-24

### Added
- **Excel Table Get Data** (#234): New `get-data` action returns table rows

### Fixed
- **Power Query Error Query Fix** (#236): Fixed spurious "Error Query" entries

## [1.3.0] - 2025-11-22

### Added
- **Chart Operations** (#229): 15 new chart actions
- **Connection Delete** (#226): New `delete` action
- **OLAP PivotTable Measures** (#217): Auto-create DAX measures

### Changed
- **PivotTable Enhancements** (#219, #220): Date/numeric grouping, calculated fields

## [1.2.0] - 2025-11-17

### Added
- **Worksheet Reordering** (#186): New `move` action

### Fixed
- **MCP Server Crash Fix** (#192): Fixed crashes with disconnected COM proxies
- **Connection Create Fix** (#190): Fixed COM dispatch error

## [1.1.0] - 2025-11-10

### Fixed
- **File Lock Fix** (#173): Fixed "file already open" errors
- **LoadTo Silent Failure Fix** (#170): LoadTo now properly fails on duplicates
- **Validation InputTitle/Message** (#167): Fixed empty values
- **Power Query Update Fix** (#140): Fixed M code merging instead of replacing
- **SetFormulas/SetValues Fix** (#199): Fixed "out of memory" error
- **Data Model Loading Fix** (#64): Fixed `set-load-to-data-model` failures
- **Power Query Persistence** (#42): Fixed load-to-data-model not persisting

### Added
- **PivotTable Discovery** (#155): Improved LLM discoverability
- **CLI Batch Support** (#152): Batch mode for bulk operations
- **Timeout Support** (#131): Configurable timeouts for all tools
- **QueryTable Support** (#129): New `excel_querytable` tool
- **Connection Create** (#127): New `create` action
- **PivotTable from Data Model** (#109): Create PivotTables from Power Pivot

### Changed
- **Numeric Column Names** (#136): Column names can now be numeric

## [1.0.0] - 2025-10-29

### Added
- Initial release of ExcelMcp
- MCP Server with 11 tools and 100+ operations
- CLI for command-line scripting
- VS Code Extension for one-click installation
- Power Query management
- Data Model / Power Pivot support
- Excel Tables and PivotTables
- Range operations with formulas
- Chart creation
- Named ranges and parameters
- VBA macro execution
- Worksheet lifecycle management
- Batch operations for performance

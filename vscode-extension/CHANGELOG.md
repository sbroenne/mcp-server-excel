# Change Log

All notable changes to the ExcelMcp VS Code extension will be documented in this file.

## [Unreleased]

- **JSON Token Optimization** (#341): Reduced token consumption for AI assistants
  - Shortened JSON property names (e.g., `sessionId` → `sid`, `sheetName` → `sn`, `address` → `addr`)
  - Split 12 monolithic tools into 21 focused tools for smaller tool schemas
  - Same 182 operations, better organized for AI assistants
  - Ranges: 4 tools (excel_range, excel_range_edit, excel_range_format, excel_range_link)
  - PivotTables: 3 tools (excel_pivottable, excel_pivottable_field, excel_pivottable_calc)
  - Tables: 2 tools (excel_table, excel_table_column)
  - Data Model: 2 tools (excel_datamodel, excel_datamodel_rel)
  - Charts: 2 tools (excel_chart, excel_chart_config)
  - Worksheets: 2 tools (excel_worksheet, excel_worksheet_style)
- **LLM Integration Testing** (#341): Added real AI agent testing using [agent-benchmark](https://github.com/mykhaliev/agent-benchmark)
  - Validates that LLMs correctly understand and use MCP tools
  - Tests incremental update patterns vs rebuild behavior
  - Ensures tool descriptions guide LLMs effectively
- **.NET 10 Upgrade**: Full compatibility with .NET 10.0
  - Updated SDK version from 8.0.416 to 10.0.100
  - All projects (Core, CLI, MCP Server) now target net10.0
  - Updated package dependencies to latest stable versions
  - GitHub Actions workflows updated to use .NET 10 SDK

## [1.4.42] - 2025-12-15

- **Power Query Rename** (#326, #327): New `rename` action for Power Query queries
- **Data Model Table Rename** (#326, #327): New `rename-table` action for Data Model tables (best-effort via Power Query)
- Performance optimizations for rename operations

## [1.4.41] - 2025-12-14

- **Power Query Data Model Fix** (#324): Fixed "0x800A03EC" error when updating Power Query in workbooks with Data Model present

## [1.4.40] - 2025-12-14

- **MCP SDK Upgrade** (#301): Upgraded ModelContextProtocol SDK from 0.4.1-preview.1 to 0.5.0-preview.1
  - Proper `isError` signaling for tool execution failures (MCP protocol compliance)
  - Deterministic exit codes (0 = success/graceful shutdown, 1 = fatal error)
  - `IAsyncDisposable` for proper async resource cleanup

## [1.4.37] - 2025-12-06

- **PivotTable Performance** (#286): Optimized `RefreshTable()` calls for faster PivotTable operations
- **Data Model Members** (#288): Added support for Data Model table members

## [1.4.36] - 2025-12-06

- **Documentation Updates** (#290): Updated tool/operation counts and lists to match code
- **SEO Fix** (#292): Fixed robots.txt sitemap URL and added Bing verification

## [1.4.35] - 2025-12-05

- **Data Model Relationships** (#278): Full support for creating, updating, and deleting relationships between Data Model tables
- **Custom Domain** (#276): Configured gh-pages for custom domain excelmcpserver.dev
  
## [1.4.34] - 2025-12-05

- **DAX Formula Locale Handling** (#281): DAX formulas with US comma separators now work correctly on European locales (German, French, etc.)

## [1.4.33] - 2025-12-04

- **Atomic Cross-File Worksheet Operations** (#273): New `copy-to-file` and `move-to-file` actions replace session-based `copy-to-workbook` and `move-to-workbook` - simpler API with no session management required

## [1.4.32] - 2025-12-04

- **OLAP PivotChart Creation** (#267): `CreateFromPivotTable` now works with OLAP/Data Model PivotTables
- **Power Query LoadToBoth Detection** (#271): Fixed incorrect detection of LoadToBoth mode

## [1.4.31] - 2025-12-04

- **Locale-Independent Number Formatting** (#263): Number and date formats now work correctly on non-US locales (German, French, etc.)

## [1.4.30] - 2025-12-03

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

## [1.4.0] - 2025-11-24

- **Excel Table Get Data** (#234): New `get-data` action returns table rows with optional `visibleOnly` filter
- **Power Query Error Query Fix** (#236): Fixed spurious "Error Query" entries when workbooks contain manual tables

## [1.3.0] - 2025-11-22

- **Chart Operations** (#229): 15 new chart actions - create charts from ranges or PivotTables, customize titles, legends, axes, styles, and series
- **Connection Delete** (#226): New `delete` action for removing data connections
- **OLAP PivotTable Measures** (#217): Auto-create DAX measures when adding value fields to Data Model PivotTables
- **PivotTable Enhancements** (#219, #220): Date/numeric grouping, calculated fields, layout control, OLAP drill operations

## [1.2.0] - 2025-11-17

- **Worksheet Reordering** (#186): New `move` action to reorder sheets within a workbook
- **MCP Server Crash Fix** (#192): Fixed crashes when closing sessions with disconnected COM proxies
- **Connection Create Fix** (#190): Fixed COM dispatch error when creating OLEDB connections in new workbooks

## [1.1.0] - 2025-11-10

- **File Lock Fix** (#173): Fixed "file already open" errors on rapid sequential operations
- **LoadTo Silent Failure Fix** (#170): LoadTo now properly fails when target worksheet already exists
- **Validation InputTitle/Message** (#167): Fixed `get-validation` returning empty InputTitle and InputMessage
- **PivotTable Discovery** (#155): Fixed LLMs not finding PivotTable functionality in MCP Server
- **CLI Batch Support** (#152): Optional batch mode for CLI commands enables bulk operations
- **Power Query Update Fix** (#140): Fixed critical bug where Update action merged M code instead of replacing
- **Numeric Column Names** (#136): Column names can now be numeric (e.g., "60" for 60 months)
- **Power Query Column Structure** (#133): Changing column structure now properly updates worksheet
- **Timeout Support** (#131): All tools now support configurable timeouts
- **QueryTable Support** (#129): New `excel_querytable` tool for legacy data import workflows
- **Connection Create** (#127): New `create` action for programmatic connection creation
- **PivotTable from Data Model** (#109): Create PivotTables from Power Pivot Data Model tables
- **SetFormulas/SetValues Fix** (#199): Fixed "out of memory" error on wide horizontal ranges
- **Data Model Loading Fix** (#64): Fixed `set-load-to-data-model` configuration failures
- **Power Query Persistence** (#42): Fixed load-to-data-model not persisting correctly

## [1.0.0] - 2025-10-29

- Initial release of ExcelMcp VS Code extension

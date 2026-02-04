# Changelog

All notable changes to ExcelMcp will be documented in this file.

This changelog covers all components:
- **MCP Server** - Model Context Protocol server for AI assistants
- **CLI** - Command-line interface for scripting and coding agents
- **VS Code Extension** - One-click installation with bundled MCP Server
- **MCPB** - Claude Desktop bundle for one-click installation

## [Unreleased]

### Added

- **CLI Daemon Improvements** (#XXX): Enhanced tray icon experience with better update management and save prompts
  - Added "Update CLI" menu option when updates are available (detects global vs local .NET tool install)
  - Added save dialog (Yes/No/Cancel) when closing individual sessions from tray
  - Added save dialog (Yes/No/Cancel) when stopping daemon with active sessions
  - Removed redundant disabled "Excel CLI Daemon" status menu entry
  - Toast notifications now mention the Update CLI menu option for easier access
  - Update command shows in confirmation dialog before execution
  - Auto-restart daemon after successful update

### Changed

- **JSON Property Names Reverted** (#417): Removed short property name mappings for better readability
  - JSON output now uses camelCase C# property names (e.g., `success`, `errorMessage`, `filePath`)
  - Removed 433 `[JsonPropertyName]` attributes from model files
  - LLMs and humans can now read JSON without consulting a mapping table

### Fixed

- **CLI Banner Cleanup**: Removed PowerShell warning from startup banner
  - Guidance moved to skill documentation (Rule 2: Use File-Based Input)
  - CLI output is now cleaner and less cluttered

- **CLI Missing Parameter Mappings** (#423): Fixed CLI commands silently ignoring user-provided values
  - ROOT CAUSE: Settings properties defined but not passed to daemon in args switch statements
  - FIX: Added missing parameter mappings for affected commands:
    - `connection set-properties`: Added `description`, `backgroundQuery`, `savePassword`, `refreshPeriod`
    - `powerquery create/load-to`: Added `targetSheet`, `targetCellAddress`
    - `chart create-*` and `move`: Added `left`, `top`, `width`, `height`
    - `table append`: Fixed to parse CSV into proper `rows` format
    - `vba run`: Added `timeoutSeconds`
  - Added pre-commit check (`check-cli-settings-usage.ps1`) to prevent future occurrences

## [1.6.5] - 2026-02-03

- **Dead Session Detection** (#414): Auto-detect and cleanup sessions when Excel process dies
  - ROOT CAUSE: `SessionManager` never checked if Excel process was alive, leaving dead sessions in dictionary
  - FIX: `GetSession()`, `GetActiveSessions()`, and `IsSessionAlive()` now check process health and auto-cleanup
  - `ExcelBatch.Execute()` validates Excel is alive before queueing operations
  - Users now get clear error: "Excel process is no longer running" instead of confusing timeouts
  - Dead sessions no longer block reopening the same file
  - Affects both CLI and MCP Server (shared `SessionManager`)

## [1.6.4] - 2026-02-03

### Fixed

- **COM Timeout with Data Model Dependencies** (#412): Fixed timeout when setting formulas/values that trigger Data Model recalculation
  - ROOT CAUSE: Excel's automatic calculation blocks COM interface during DAX recalculation
  - FIX: Temporarily disable calculation mode (xlCalculationManual) during write operations
  - Affected methods: `SetFormulas`, `SetValues`, `Table.Append`, `NamedRange.Write`
  - Formulas like `=INDEX(KPIs[Total_ACR],1)` now work without "The operation was canceled" error

## [1.6.3] - 2026-02-03

### Documentation

- **M Code Identifier Quoting** (#407): Added guidance for special characters in Power Query identifiers
- **PowerQuery Eval-First Workflow** (#405): Updated documentation with eval-first pattern
- **CLI Command Name Fix** (#403): Fixed CLI command name in agent skills installation docs

## [1.6.2] - 2026-02-02

### Fixed

- **Power Query Refresh Error Propagation** (#399): Fixed bug where `refresh` action returned `success: true` even when Power Query had formula errors
  - ROOT CAUSE: `Connection.Refresh()` silently swallows errors for worksheet queries (InModel=false)
  - FIX: Now uses `QueryTable.Refresh(false)` for worksheet queries which properly throws errors
  - Data Model queries (InModel=true) continue using `Connection.Refresh()` which does throw errors
  - Errors now surface clearly: `"[Expression.Error] The name 'Source' wasn't recognized..."`

- **Table Create Auto-Expand from Single Cell**: Fixed issue where `table create --range A1` created single-cell table
  - ROOT CAUSE: Excel's `ListObjects.Add()` doesn't auto-expand from a single cell
  - FIX: Now uses `Range.CurrentRegion` when single cell provided, capturing all contiguous data
  - Prevents Data Model issues where tables only contain header column

### Added

- **Power Query Evaluate** (#400): New `evaluate` action to execute M code directly and return results
  - Execute arbitrary M code without creating a permanent query
  - Returns tabular results (columns, rows) in JSON format
  - Automatically cleans up temporary query and worksheet
  - Errors propagate properly (e.g., invalid M syntax throws with error message)
  - Example: `excelcli powerquery evaluate --file data.xlsx --mcode "let Source = #table({\"Name\",...})"`

- **MCP Power Query mCodeFile Parameter**: Read M code from file instead of inline string
  - New `mCodeFile` parameter on `excel_powerquery` tool for `create`, `update`, `evaluate` actions
  - Avoids JSON escaping issues with complex M code containing special characters
  - File takes precedence if both `mCode` and `mCodeFile` provided

- **MCP VBA vbaCodeFile Parameter**: Read VBA code from file instead of inline string
  - New `vbaCodeFile` parameter on `excel_vba` tool for `create-module`, `update-module` actions
  - Handles VBA code with quotes and special characters cleanly
  - File takes precedence if both `vbaCode` and `vbaCodeFile` provided

## [1.6.1] - 2026-02-01

### Fixed

- **CLI PackAsTool Workaround** (#396): Fixed CLI packaging issue with net10.0-windows target
- **CI Duplicate Paths** (#394): Removed duplicate paths key in build workflow

## [1.6.0] - 2026-02-01

### Fixed

- **MCPB Skills Key** (#392): Removed unsupported 'skills' key from manifest
- **Data Model MSOLAP Error** (#391): Better error message when MSOLAP provider is missing

## [1.5.14] - 2025-02-01

### Added

#### CLI Redesign (Breaking Change)
- **Complete CLI Rewrite** (#387): Redesigned CLI for coding agents and scripting - **NOT backwards compatible**
  - 14 unified command categories with 210 operations matching MCP Server
  - All commands now use `--session` parameter (was positional in some commands)
  - Comprehensive `--help` descriptions on all commands synced with MCP tool descriptions
  - All `--file` parameters support both new file creation and existing files
  - New `excelcli list-actions` command to discover all available operations
  - Exit code standardization (0=success, 1=error, 2=validation)

- **Quiet Mode**: `-q`/`--quiet` flag suppresses banner for agent-friendly JSON-only output
  - Auto-detects piped/redirected stdout and suppresses banner automatically

- **Version Check**: `excelcli version --check` queries NuGet to show if update available

- **Session Close --save**: Single `--save` flag for atomic save-and-close workflow
  - Replaces separate save + close sequence for cleaner scripting

- **CLI Action Coverage Pre-commit Check**: New `check-cli-action-coverage.ps1` script
  - Ensures CLI switch statements cover ALL action strings from ActionExtensions.cs
  - Prevents "action not handled" bugs from reaching production
  - Validates 210 operations across 21 CLI commands

#### MCP Server Enhancements  
- **Session Operation Timeout** (#388): Configurable timeout prevents infinite hangs
  - New `timeoutSeconds` parameter on `excel_file(open)` and `excel_file(create)` actions
  - Default: 300 seconds (5 minutes), configurable range: 10-3600 seconds
  - Applies to ALL operations within session; exceeding timeout throws `TimeoutException`

- **Create Action** (#385): Renamed `create-and-open` to simpler `create` action
  - Single-action file creation and session opening
  - Performance: ~3.8 seconds (vs ~7-8 seconds with separate create+open)

- **PowerQuery Unload Action**: New `unload` action removes data from all load destinations
  - Keeps query definition intact while clearing worksheet/model data

#### Testing & Quality
- **LLM Integration Tests**: Comprehensive agent-benchmark test suite for CLI
  - 9 test scenarios covering all major Excel operations
  - Chart positioning, PivotTable layout, Power Query, slicers, tables, ranges
  - Financial report automation workflow tests

- **Agent Skills**: New structured skills documentation for AI assistants
  - `skills/excel-cli/` - CLI-specific skill with commands reference
  - `skills/excel-mcp/` - MCP Server skill with tools reference
  - `skills/shared/` - Shared workflows, anti-patterns, behavioral rules

### Fixed
- **Calculated Field Bug**: Fixed PivotTable calculated field creation error
- **COM Diagnostics**: Improved error reporting for COM object lifecycle issues

### Changed
- CLI timeout option uses `--timeout <seconds>` (was `--timeout-seconds`)
- All CLI commands now require explicit `--session` parameter

## [1.5.13] - 2025-01-24

### Added
- **Chart Formatting** (#384): Enhanced chart formatting capabilities
  - **Data Labels**: Configure label position and visibility (showValue, showCategory, showPercentage, etc.)
  - **Axis Scale**: Get/set axis scale properties (min, max, units, auto-scale flags)
  - **Gridlines**: Control major/minor gridlines visibility on chart axes
  - **Series Markers**: Configure marker style, size, and colors for data series
  - 8 new operations bringing total chart operations to 22

- **Chart Trendlines** (#386): Statistical analysis and forecasting for chart series
  - **Add Trendline**: Linear, Exponential, Logarithmic, Polynomial, Power, Moving Average
  - **List Trendlines**: View all trendlines on a series
  - **Delete Trendline**: Remove trendline by index
  - **Configure Trendline**: Forward/backward forecasting, display equation and R² value
  - 4 new operations bringing total chart operations to 26

## [1.5.11] - 2025-01-22

### Added
- Added Agent Skill to all artifacts

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

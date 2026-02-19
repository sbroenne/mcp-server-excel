# Changelog

All notable changes to ExcelMcp will be documented in this file.

This changelog covers all components:
- **MCP Server** - Model Context Protocol server for AI assistants
- **CLI** - Command-line interface for scripting and coding agents
- **VS Code Extension** - One-click installation with bundled MCP Server
- **MCPB** - Claude Desktop bundle for one-click installation

## [Unreleased]

### Fixed

- **COM Apartment Boundary in SaveWorkbook** (#482): Removed `Task.Run(() => workbook.Save())` in `ExcelShutdownService` — this marshalled the COM call from the STA thread to an MTA thread-pool thread, which is incorrect and fragile in .NET 8+. Save is now called directly on the STA thread, which is always the case inside `ExcelBatch.Execute()`.
- **Wrong-Process Force-Kill from Fallback PID** (#482): Removed the "newest EXCEL.EXE process" fallback PID detection in `ExcelBatch`. When the `Hwnd` path fails, force-kill is now disabled with a warning rather than risking killing an unrelated Excel workbook the user has open.
- **Redundant `Thread.Sleep` in Dispose** (#482): Removed 100 ms `Thread.Sleep` from `ExcelBatch.Dispose()`. The preceding `_shutdownCts.Cancel()` call immediately wakes the STA thread from `WaitToReadAsync`, making the sleep redundant and adding unnecessary latency.
- **Exception Type Lost in Service Error Responses** (#482): `ExcelMcpService` top-level `catch` blocks now return `"{ExType}: {ex.Message}"` instead of just `ex.Message`, making unexpected failures distinguishable without a full stack trace.
- **COM Timeout Hang** — ExcelBatch now force-kills Excel process on timeout instead of hanging indefinitely on `WaitForSingleObject`; ExcelMcpService catches `TimeoutException` to prevent unhandled exceptions
- **FileSystemWatcher CPU Spin** — Disabled `IConfiguration` reload-on-change in MCP Server to prevent 85%+ CPU usage from `FileSystemWatcher` polling
- **Process Handle Leak** — Fixed `Process` object not being disposed in `ExcelBatch.ForceKillExcelProcess()`
- **Configuration Sources Cleared** — Re-add environment variables and command-line args after clearing config sources (were accidentally removed)
- **Source Generator Type Aggregation** — Fixed nullable type upgrade logic in `ServiceInfoExtractor` that could lose type information across partial interfaces
- **Chart Trendline Parameter Name** — Renamed `type` → `trendlineType` in `IChartConfigCommands` to avoid COM parameter ambiguity
- **Chart Style Error Message** — Improved `SetStyle` error message to show valid range when `styleId` is out of bounds
- **Chart InvalidOperationException** — Added catch for `InvalidOperationException` in chart appearance commands

### Changed

- **Chart Test Performance** — Refactored 80 chart tests to share a single pre-populated fixture file via `File.Copy()` instead of creating individual files via COM, eliminating ~74 redundant Excel sessions

### Added

- **Window Management Tool** (#470): New `window` tool with 9 operations to control Excel window visibility, position, state, and status bar — enabling "Agent Mode" where users watch AI work in Excel
  - `show` / `hide` — Toggle Excel visibility (syncs with session metadata)
  - `bring-to-front` — Bring Excel to foreground
  - `get-info` — Query window state (visibility, position, size, foreground status)
  - `set-state` — Set normal / minimized / maximized
  - `set-position` — Set window left, top, width, height
  - `arrange` — Preset layouts: left-half, right-half, top-half, bottom-half, center, full-screen
  - `set-status-bar` — Display live operation status text in Excel's status bar
  - `clear-status-bar` — Restore default status bar text
  - MCP Server proactively asks users about showing Excel for visual tasks (charts, dashboards)
  - Agent Mode, Presentation Mode, and Debug Mode workflow guidance
- **CLI `--output` flag** for all commands: Save command output directly to a file. Screenshot commands automatically save decoded PNG images instead of base64 JSON
- **CLI Batch Mode** (#463): New `excelcli batch` command executes multiple CLI commands from a JSON file in a single process launch
  - Session auto-capture from `session.open`/`session.create`, auto-clear on `session.close`
  - NDJSON output for machine-readable results
  - `--stop-on-error` flag to halt on first failure (default: continue all)

### Fixed

- **Screenshot reliability**: Screenshots now work reliably regardless of whether Excel is visible or hidden. Added automatic retry for transient capture failures
- **CLI `--help` crash** (#463): Fixed Spectre.Console markup crash when parameter descriptions contain `[`/`]` characters (e.g., `[A1 notation]`)
- **Source generator tool filtering**: Fixed `mcpTool ?? "unknown"` fallback; added `HasMcpToolAttribute` to correctly filter MCP-only tools
- **Skills docs parameter names**: Fixed wrong CLI parameter names in `conditionalformat.md` and `slicer.md` reference files
- **Auto-save on shutdown**: Sessions are now auto-saved before closing when MCP server exits or client disconnects, preventing silent data loss from session timeouts
- **Session creation resilience**: Added retry logic (Polly) for transient COM failures (`CO_E_SERVER_EXEC_FAILURE`, `RPC_E_CALL_FAILED`) during Excel process startup under resource constraints

## [1.7.2] - 2026-02-15

### Added

- **In-Process Service Architecture** (#454): MCP Server and CLI each host ExcelMCP Service in-process instead of sharing a separate service process
  - Eliminates service discovery failures (especially NuGet tool installs) and cross-process coordination

- **Separate CLI NuGet Package** (#452): CLI published as `Sbroenne.ExcelMcp.CLI` alongside MCP Server
  - Service version negotiation: client validates exact version match with running service on connect

### Fixed

- **Build Workflow Path** (#455): Fixed target framework path (`net10.0` → `net10.0-windows`) and formatting errors in build workflow

## [1.7.1] - 2026-02-09

### Fixed

- **Release Workflow** (#451): Moved all external publishing steps after builds succeed to prevent partial releases

## [1.6.10] - 2026-02-06

### ⚠️ BREAKING CHANGES

**See [BREAKING-CHANGES.md](docs/BREAKING-CHANGES.md) for complete migration guide.**
LLMs pick up these changes automatically via `tools/list` (MCP) and `--help` (CLI).

- **Tool Names Simplified**: Removed `excel_` prefix from all 23 MCP tool names (e.g., `excel_range` → `range`, `excel_file` → `file`). Titles also shortened (e.g., `"Chart Operations"`). VS Code extension server name → `excel-mcp`.

### Added

- **CLI Code Generation** (#433): CLI commands auto-generated from Core via Roslyn source generators — guarantees 1:1 MCP/CLI parity
- **Calculation Mode Control** (#430): New `calculation_mode` tool/CLI command (automatic, manual, semi-automatic modes; workbook/sheet/range scopes)
- **Installation via npx** (#449): Added `npx add-mcp` as primary installation method in docs

### Changed

- **MCP Prompt Reduction** (#442): Reduced prompts from 7 to 4 with ~76% content reduction; removed `excel_` prefix from prompt names
- **VS Code Extension**: Self-contained publishing (no .NET runtime needed), CLI removed from extension, skills use `chatSkills` contribution point
- **LLM Tests** (#446): Migrated to pytest-aitest v0.3.x from PyPI with unified MCP/CLI test suite
- **Release Workflow** (#443): Switched to workflow_dispatch with version bump UI; added stale issue workflow
- **Terminology**: "Daemon" → "ExcelMCP Service" throughout docs
- **MCP SKILL template** (#448): Added Workflow Checklist table for quick reference (open → create → write → format → save)
- **CLI SKILL template** (#448): Added "List Parameters Use JSON Arrays" to Common Pitfalls section
- **Slicer reference doc**: Added CLI JSON Array Quoting section with PowerShell escaping examples
- **MCPB**: Removed agent skills from Claude Desktop bundle

### Fixed

- **MCP Server Release Path** (#450): Corrected package path to `net10.0-windows`
- **Broken Emoji Characters**: Fixed corrupted emoji in README files

### Removed

- **Glama.ai Support**: Removed Docker-based deployment (`Dockerfile`, `glama.json`, `.dockerignore`, docs)

## [1.6.9] - 2026-02-04

### Added

- **CLI Daemon Improvements**: Enhanced tray icon experience with better update management and save prompts
  - Added "Update CLI" menu option when updates are available (detects global vs local .NET tool install)
  - Added save dialog (Yes/No/Cancel) when closing individual sessions from tray
  - Added save dialog (Yes/No/Cancel) when stopping daemon with active sessions
  - Removed redundant disabled "Excel CLI Daemon" status menu entry
  - Toast notifications now mention the Update CLI menu option for easier access
  - Update command shows in confirmation dialog before execution
  - Auto-restart daemon after successful update

### Fixed

- **PivotTable RPC Disconnection** (#426): Fixed "RPC server is unavailable (0x800706BA)" error during rapid OLAP PivotTable field operations
  - ROOT CAUSE: `RefreshTable()` called after each field operation triggered synchronous Analysis Services queries
  - FIX: Removed RefreshTable() from field manipulation methods (AddRowField, AddColumnField, AddFilterField, RemoveField, SetFieldFunction)
  - Field changes now take effect immediately without blocking AS queries
  - Call `pivottable(refresh)` explicitly to update visual display after configuring fields
  - Applies to both OLAP (Data Model) and regular PivotTables for consistency

## [1.6.8] - 2026-02-03

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
  - New `mCodeFile` parameter on `powerquery` tool for `create`, `update`, `evaluate` actions
  - Avoids JSON escaping issues with complex M code containing special characters
  - File takes precedence if both `mCode` and `mCodeFile` provided

- **MCP VBA vbaCodeFile Parameter**: Read VBA code from file instead of inline string
  - New `vbaCodeFile` parameter on `vba` tool for `create-module`, `update-module` actions
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
  - New `timeoutSeconds` parameter on `file(open)` and `file(create)` actions
  - Default: 300 seconds (5 minutes), configurable range: 10-3600 seconds
  - Applies to ALL operations within session; exceeding timeout throws `TimeoutException`

- **Create Action** (#385): Renamed `create-and-open` to simpler `create` action
  - Single-action file creation and session opening
  - Performance: ~3.8 seconds (vs ~7-8 seconds with separate create+open)

- **PowerQuery Unload Action**: New `unload` action removes data from all load destinations
  - Keeps query definition intact while clearing worksheet/model data

#### Testing & Quality
- **LLM Integration Tests**: Comprehensive pytest-aitest test suite for CLI
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
- **PivotTable & Table Slicers** (#363): New `slicer` tool for interactive filtering
  - **PivotTable Slicers**: Create, list, filter, and delete slicers for PivotTable fields
  - **Table Slicers**: Create, list, filter, and delete slicers for Excel Table columns
  - 8 new operations for interactive data filtering

## [1.5.5] - 2025-01-19

### Added
- **DMV Query Execution** (#353): Query Data Model metadata using Dynamic Management Views
  - New `execute-dmv` action on `datamodel` tool
  - Query TMSCHEMA_MEASURES, TMSCHEMA_RELATIONSHIPS, DISCOVER_CALC_DEPENDENCY, etc.

## [1.5.4] - 2025-01-19

### Added
- **DAX EVALUATE Query Execution** (#356): Execute DAX queries against the Data Model
  - New `evaluate` action on `datamodel` tool for ad-hoc DAX queries
- **DAX-Backed Excel Tables** (#356): Create worksheet tables populated by DAX queries
  - New `create-from-dax`, `update-dax`, `get-dax` actions

## [1.5.0] - 2025-01-10

### Changed
- **Tool Reorganization** (#341): Split 12 monolithic tools into 21 focused tools
  - 186 operations total, better organized for AI assistants
  - Ranges: 4 tools (range, range_edit, range_format, range_link)
  - PivotTables: 3 tools (pivottable, pivottable_field, pivottable_calc)
  - Tables: 2 tools (table, table_column)
  - Data Model: 2 tools (datamodel, datamodel_rel)
  - Charts: 2 tools (chart, chart_config)
  - Worksheets: 2 tools (worksheet, worksheet_style)

### Added
- **LLM Integration Testing** (#341): Real AI agent testing using [pytest-aitest](https://github.com/sbroenne/pytest-aitest)

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

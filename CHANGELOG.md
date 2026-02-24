# Changelog

All notable changes to ExcelMcp will be documented in this file.

## [Unreleased]

### Added
- **IRM/AIP-protected file support** in `file(action='open')`: Files protected with Azure Information Protection (AIP/IRM) are automatically detected via OLE2 Compound Document header signature (`D0 CF 11 E0 A1 B1 1A E1`). They are opened read-only with Excel forced visible so the Windows IRM credential prompt can appear — no extra parameters needed.
- **`isIrmProtected` field in `file(action='test')` response**: Reports whether a file uses the OLE2/IRM format before attempting to open it, enabling agents to pre-flight check and set expectations (IRM files are always read-only).

### Changed
- **MCP SDK upgraded** from `ModelContextProtocol` 0.8.0-preview.1 to 0.9.0-preview.2
  - Updated `ImageContentBlock.Data` from `string` to `ReadOnlyMemory<byte>` (binary data API change)
- CLI daemon IPC migrated from custom protocol to StreamJsonRpc — improved reliability and error handling
- Enhanced daemon security with SID validation and `CurrentUserOnly` pipe access

### Fixed
- **Chart `create-from-range`** (#512): sheet names with spaces or special characters now quoted in source data reference — fixes COM error when source sheet name contains spaces
- **PivotTable `create-from-range`** (#512): same quoting fix for pivot cache source data reference
- **Worksheet `rename`** (#513): `target_name` parameter description now mentions rename action, improving LLM discoverability
- Standardized all data operation timeouts to 30 minutes via `ComInteropConstants.DataOperationTimeout`
  - Power Query `load-to` increased from 5 min to 30 min
  - Connection `refresh` and `load-to` increased from 5 min to 30 min
  - Power Query `refresh`/`refresh-all` now use the same constant (was inline 30 min)
  - `ExcelBatch.Execute()` no longer double-caps timeout when caller provides a cancellation token
- Corrected stale documentation claiming 60-600 second range restriction on refresh timeout (no such validation exists)

## [1.8.12] - 2026-02-22

### Fixed
- Power Query `refresh` and `refresh-all` no longer crash when very large timeout values are provided

## [1.8.11] - 2026-02-22

### Fixed
- `--timeout` parameter on `powerquery refresh-all` now works end-to-end (was silently ignored)
- CLI pipe timeout no longer races with operation timeout on long-running refreshes

## [1.8.10] - 2026-02-22

### Fixed
- CLI daemon no longer spawns duplicate instances during long-running operations (e.g. `powerquery refresh-all`)

## [1.8.9] - 2026-02-22

### Fixed
- `powerquery load-to` now correctly transitions between all load destinations (e.g. worksheet → connection-only, data-model → worksheet)
- `powerquery get-load-config` no longer crashes on connection-only queries
- CLI banner no longer pollutes JSON output when piped (e.g. `excelcli ... | ConvertFrom-Json`)

## [1.8.8] - 2026-02-22

### Fixed
- `powerquery refresh` no longer crashes when `--timeout` is omitted (uses 5-minute default)

## [1.8.7] - 2026-02-21

### Fixed
- `table add-to-data-model` now handles column names containing brackets — new `stripBracketColumnNames` parameter
- `powerquery load-to data-model` now actually loads data into the Data Model (was silently no-op)
- `chartconfig set-data-labels` gives clear error when using bar-only label positions on Line charts
- `rangeformat format-range` — corrected `borderStyle` valid values in documentation
- `rangeformat format-range` now accepts `middle` as alias for `center` vertical alignment

## [1.8.6] - 2026-02-21

### Fixed
- Workbooks with connections, Power Query, or Data Model no longer crash on open in NuGet tool installs

## [1.8.3] - 2026-02-19

### Fixed
- Operations that trigger Excel recalculation (conditional formatting on formula cells, PivotTable refresh, Power Query refresh) no longer hang
- Session no longer becomes permanently stuck after cancelling a tool call
- After a timeout, subsequent operations now fail immediately instead of waiting for another full timeout

## [1.8.2] - 2026-02-19

### Fixed
- Excel process no longer hangs on timeout — forcefully recovered instead
- Fixed 85%+ CPU usage caused by configuration file watcher
- Session save on workbook close now works reliably
- Improved error messages with exception type information
- Fixed errors when saving workbooks during shutdown

## [1.8.0] - 2026-02-17

### Added
- **Screenshot quality parameter** — `High`/`Medium`/`Low` quality settings (default: `Medium`)
- **Window Management** — new `window` tool with 9 operations: show/hide Excel, arrange windows, set status bar text
- **CLI `--output` flag** — save command output directly to a file
- **CLI Batch Mode** — execute multiple CLI commands from a JSON file in a single launch

### Fixed
- Screenshots now work reliably regardless of Excel visibility
- Worksheet `copy-to-file` and `move-to-file` no longer require a session parameter
- CLI `--help` no longer crashes on commands with bracket characters in descriptions
- Sessions auto-save when MCP server exits or client disconnects

## [1.7.7] - 2026-02-17

### Added
- Chart collision detection and auto-positioning with `targetRange` parameter

## [1.7.6] - 2026-02-16

### Added
- **CLI Batch Mode** — `excelcli batch` executes multiple commands from a JSON file

## [1.7.5] - 2026-02-16

### Fixed
- Fixed 100% CPU usage in background message pump loop

## [1.7.4] - 2026-02-16

### Fixed
- Fixed crash when accessing non-OLEDB connection types

## [1.7.2] - 2026-02-15

### Added
- **In-process service architecture** — MCP Server and CLI each host the service internally, eliminating service discovery failures
- **CLI NuGet package** — CLI now published as `Sbroenne.ExcelMcp.CLI`

## [1.7.1] - 2026-02-15

### Fixed
- Release workflow no longer publishes partial releases on build failures

## [1.6.10] - 2026-02-15

### ⚠️ BREAKING CHANGES

**See [BREAKING-CHANGES.md](docs/BREAKING-CHANGES.md) for migration guide.**

- Tool names simplified: removed `excel_` prefix (e.g. `excel_range` → `range`)

### Added
- **Calculation Mode Control** — new `calculation_mode` tool (automatic, manual, semi-automatic)
- CLI commands now auto-generated from Core — guaranteed 1:1 MCP/CLI parity
- Added `npx add-mcp` installation method

### Removed
- Docker/Glama.ai deployment support

## [1.6.9] - 2026-02-04

### Added
- CLI tray icon: "Update CLI" menu option, save prompts when closing sessions

### Fixed
- PivotTable field operations no longer cause "RPC server is unavailable" errors on Data Model PivotTables

## [1.6.8] - 2026-02-03

### Fixed
- CLI commands no longer silently ignore user-provided parameter values for `connection`, `powerquery`, `chart`, `table`, and `vba` commands
- JSON output uses readable property names (reverted short names)

## [1.6.5] - 2026-02-03

### Fixed
- Dead Excel sessions are now auto-detected and cleaned up — no more confusing timeouts when Excel process dies

## [1.6.4] - 2026-02-03

### Fixed
- Setting formulas/values that trigger Data Model recalculation no longer times out

## [1.6.3] - 2026-02-03

### Fixed
- Documentation improvements for Power Query M code and CLI command names

## [1.6.2] - 2026-02-02

### Added
- **Power Query `evaluate`** — execute M code directly and return results without creating a permanent query
- **`mCodeFile` parameter** — read M code from file instead of inline string (avoids escaping issues)
- **`vbaCodeFile` parameter** — read VBA code from file instead of inline string

### Fixed
- `powerquery refresh` now correctly reports errors instead of returning success on formula errors
- `table create` from a single cell now auto-expands to the full contiguous data range

## [1.6.0] - 2026-02-01

### Added
- **Complete CLI Rewrite** — 14 command categories with 210 operations matching MCP Server
- **Quiet Mode** — `-q` flag for clean JSON-only output
- **Session timeout** — configurable `timeoutSeconds` on `file(open)` and `file(create)`
- **PowerQuery `unload`** — remove data from load destinations while keeping query definition

### Fixed
- Better error message when MSOLAP provider is missing for Data Model operations

## [1.5.13] - 2026-01-26

### Added
- **Chart Formatting** — data labels, axis scale, gridlines, series markers (8 new operations)
- **Chart Trendlines** — add, list, delete, and configure trendlines (4 new operations)

## [1.5.11] - 2026-01-22

### Added
- Agent Skills included in all artifacts
- MCPB bundle now includes LICENSE and CHANGELOG.md

## [1.5.8] - 2026-01-20

### Added
- Available as a Claude Desktop MCPB Extension

## [1.5.6] - 2026-01-20

### Added
- **Slicers** — create, list, filter, and delete slicers for PivotTables and Tables (8 new operations)

## [1.5.5] - 2026-01-19

### Added
- **DMV Queries** — query Data Model metadata via Dynamic Management Views

## [1.5.4] - 2026-01-19

### Added
- **DAX Queries** — execute DAX EVALUATE queries against the Data Model
- **DAX-backed Tables** — create worksheet tables populated by DAX queries

## [1.5.0] - 2026-01-09

### Changed
- **Tool Reorganization** — split 12 tools into 21 focused tools (186 operations total)
- **.NET 10 Upgrade** — requires .NET 10.0

## [1.4.42] - 2025-12-15

### Added
- Power Query `rename` and Data Model `rename-table` actions

## [1.4.41] - 2025-12-14

### Fixed
- Fixed error when updating Power Query in workbooks with Data Model

## [1.4.40] - 2025-12-14

### Changed
- Upgraded MCP SDK to 0.5.0 with proper error signaling

## [1.4.37] - 2025-12-06

### Added
- Data Model table members support
- PivotTable refresh performance improvements

## [1.4.35] - 2025-12-05

### Added
- **Data Model Relationships** — create, update, and delete relationships

## [1.4.34] - 2025-12-05

### Fixed
- DAX formulas now work on European locales (comma as decimal separator)

## [1.4.33] - 2025-12-04

### Added
- Cross-file worksheet operations: `copy-to-file` and `move-to-file`

## [1.4.32] - 2025-12-04

### Fixed
- Charts can now be created from Data Model PivotTables

## [1.4.31] - 2025-12-04

### Fixed
- Number and date formats now work correctly on non-US locales

## [1.4.30] - 2025-12-03

### Added
- `showExcel: true` parameter to watch AI work in Excel in real time

### Fixed
- Fixed errors when adding value fields to Data Model PivotTables

## [1.4.25] - 2025-12-01

### Changed
- VS Code extension is 89% smaller (framework-dependent deployment)

## [1.4.24] - 2025-12-01

### Added
- PivotTable grand totals, date/numeric grouping, calculated fields, layout control (172 total operations)

### Fixed
- MCP Server no longer stops due to network errors

## [1.4.0] - 2025-11-24

### Added
- Table `get-data` action

### Fixed
- Power Query no longer creates spurious "Error Query" entries

## [1.3.0] - 2025-11-22

### Added
- 15 chart operations
- Connection `delete` action
- OLAP PivotTable measures

## [1.2.0] - 2025-11-17

### Added
- Worksheet `move` action for reordering sheets

### Fixed
- Fixed crashes with disconnected COM proxies

## [1.1.0] - 2025-11-10

### Added
- PivotTables from Data Model, connection `create`, configurable timeouts, `querytable` tool, CLI batch mode

### Fixed
- File lock errors, Power Query update overwriting, formula/value "out of memory" error, Data Model loading

## [1.0.0] - 2025-10-29

### Added
- Initial release — MCP Server with 11 tools and 100+ operations
- CLI for command-line scripting
- VS Code Extension
- Power Query, Data Model, Tables, PivotTables, Ranges, Charts, Named Ranges, VBA, Worksheets

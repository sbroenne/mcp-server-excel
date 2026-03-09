# Changelog

This changelog is for end users. It focuses on visible product changes and uses the tagged release history as the source of truth.

## [Unreleased]

- Added an Awesome Copilot plugin bundle for easier marketplace distribution.
- Improved Excel session stability, shutdown handling, and refresh reliability.
- Fixed CLI daemon autostart so a stuck existing daemon now fails fast with a clear recovery error instead of timing out for a long period and reporting a misleading startup timeout.

## [1.8.25] - 2026-03-03

- Fixed formulas inside Excel Tables so they no longer get unwanted `@` implicit-intersection markers.

## [1.8.24] - 2026-03-02

- Added formula validation, smarter formula detection, and better handling of IRM-protected workbooks.

## [1.8.23] - 2026-02-28

- Fixed Power Query loads to named tables so `CurrentWorkbook()` references remain reliable.

## [1.8.22] - 2026-02-28

- Added progress reporting and better cancellation for Data Model and Power Query refresh operations.

## [1.8.21] - 2026-02-27

- Fixed CPU spin and session hangs during Power Query and Data Model refresh.

## [1.8.20] - 2026-02-27

- Improved cancellation and timeout handling for Connection and Power Query refresh operations.

## [1.8.19] - 2026-02-26

- Fixed CLI and MCP Server version reporting.

## [1.8.18] - 2026-02-26

- Fixed `table append` JSON value handling and added stdin support for `--values -` and `--rows -`.

## [1.8.17] - 2026-02-25

- Added installable npm skill packages for `excel-mcp` and `excel-cli`.

## [1.8.16] - 2026-02-24

- Added support for IRM/AIP-protected Excel files.

## [1.8.15] - 2026-02-23

- Standardized long-running data operation timeouts.

## [1.8.14] - 2026-02-23

- Switched CLI daemon IPC to StreamJsonRpc for more reliable communication.

## [1.8.13] - 2026-02-23

- Fixed chart and PivotTable source references when sheet names need quoting.

## [1.8.12] - 2026-02-22

- Prevented invalid oversized timeout values and fixed timeout handling for refresh-all workflows.

## [1.8.11] - 2026-02-22

- Improved CLI daemon stability during heavy refresh workloads.

## [1.8.10] - 2026-02-22

- Fixed Power Query `load-to` state handling and improved `get-load-config` accuracy.

## [1.8.9] - 2026-02-22

- Restored a safer default timeout for Power Query refresh.

## [1.8.8] - 2026-02-22

- Fixed Data Model issues with bracketed column names and improved Power Query load-to-data-model behavior.

## [1.8.7] - 2026-02-21

- Improved `office.dll` resolution for NuGet installs and Office 365 environments.

## [1.8.6] - 2026-02-21

- Improved compatibility around chart labels, screenshot docs, and Power Query load-to workflows.

## [1.8.5] - 2026-02-20

- Fixed `office.dll` startup crashes introduced during the Office interop migration.

## [1.8.4] - 2026-02-20

- Migrated core Excel interop to strongly typed Office PIAs.

## [1.8.3] - 2026-02-19

- Fixed missing worksheet tool descriptions.

## [1.8.2] - 2026-02-19

- Added live Excel "Agent Mode" window controls and status-bar feedback.

## [1.8.1] - 2026-02-17

- Fixed missing tool descriptions in the worksheet tool.

## [1.8.0] - 2026-02-17

- Added atomic worksheet file operations and improved screenshot rendering.

## [1.7.9] - 2026-02-17

- Improved COM timeout handling, CPU usage, and general runtime stability.

## [1.7.8] - 2026-02-17

- Added chart auto-positioning, collision detection, and target-range placement.

## [1.7.7] - 2026-02-17

- Improved skill packaging and auto-loading for ExcelMcp skills.

## [1.7.6] - 2026-02-16

- Added CLI batch mode for bulk workflows.

## [1.7.5] - 2026-02-16

- Updated installation and packaging flow for the current MCP and CLI setup.

## [1.7.4] - 2026-02-16

- Fixed STA CPU spin and improved connection handling for non-OLEDB sources.

## [1.7.3] - 2026-02-16

- Improved screenshot reliability and refreshed generated MCP guidance.

## [1.7.2] - 2026-02-15

- Moved MCP Server and CLI to the in-process ExcelMcp service architecture.
- Published the CLI as a separate NuGet package.

## [1.7.1] - 2026-02-09

- Fixed release publishing order to avoid partial releases.

## [1.6.10] - 2026-02-15

- Breaking: removed the `excel_` prefix from MCP tool names and simplified installation.
- Added generated CLI parity, calculation mode control, and a leaner MCP prompt set.

## [1.6.9] - 2026-02-04

- Improved the CLI tray app with save prompts and update actions.
- Fixed RPC errors during rapid PivotTable field updates.

## [1.6.8] - 2026-02-03

- Restored readable camelCase JSON output.
- Fixed missing CLI parameter mappings.

## [1.6.7] - 2026-02-03

- Added hex color input support for `set-tab-color` in the CLI.

## [1.6.6] - 2026-02-03

- Reverted abbreviated JSON property names for easier reading.

## [1.6.5] - 2026-02-03

- Added dead-session detection and cleanup when Excel exits unexpectedly.

## [1.6.4] - 2026-02-03

- Added version checking for the CLI daemon and MCP Server.

## [1.6.3] - 2026-02-03

- Added Power Query identifier quoting guidance and command-name documentation fixes.

## [1.6.2] - 2026-02-02

- Added Power Query `evaluate`.
- Improved refresh error reporting and added file-based M code and VBA code inputs.

## [1.6.1] - 2026-02-02

- Fixed CLI packaging for `net10.0-windows`.

## [1.6.0] - 2026-02-02

- Fixed MCPB packaging and improved missing-provider errors for the Data Model.

## [1.5.14] - 2026-02-01

- Breaking: redesigned the CLI for coding agents, with explicit sessions, quiet mode, version checks, and better help text.
- Added session timeouts, `create`, `unload`, and packaged agent skills.

## [1.5.13] - 2026-01-26

- Added chart positioning improvements.

## [1.5.12] - 2026-01-23

- Added chart formatting features such as labels, markers, axes, and gridlines.

## [1.5.11] - 2026-01-22

- Added packaged agent skills and improved release artifacts.

## [1.5.10] - 2026-01-22

- Added chart axis number-format actions.

## [1.5.9] - 2026-01-20

- Documentation maintenance release.

## [1.5.8] - 2026-01-20

- Added Claude Desktop MCPB distribution.

## [1.5.7] - 2026-01-20

- Added `layoutStyle` for PivotTable creation.

## [1.5.6] - 2026-01-20

- Added slicer support for PivotTables and Excel Tables.

## [1.5.5] - 2026-01-19

- Added Data Model DMV query execution.

## [1.5.4] - 2026-01-19

- Added DAX `EVALUATE` support and DAX-backed worksheet tables.

## [1.5.3] - 2026-01-18

- Added automatic Power Query M-code formatting.

## [1.5.2] - 2026-01-12

- Documentation alignment for the .NET 10 transition.

## [1.5.1] - 2026-01-10

- Restored descriptive MCP parameter names.

## [1.5.0] - 2026-01-09

- Reorganized the product into smaller focused tools.
- Upgraded to .NET 10 and added LLM integration testing.

## [1.4.42] - 2025-12-15

- Rolled back the VS Code status bar MCP monitor.

## [1.4.41] - 2025-12-14

- Improved Data Model guidance and shortened the default timeout.

## [1.4.40] - 2025-12-14

- Added a VS Code status bar MCP monitor with session visibility.

## [1.4.37] - 2025-12-06

- Documentation update to align tool and operation counts with the code.

## [1.4.36] - 2025-12-06

- Improved PivotTable refresh performance.

## [1.4.35] - 2025-12-05

- Added better DAX formula handling for European locales.

## [1.4.34] - 2025-12-05

- Added read support for individual Data Model relationships.

## [1.4.33] - 2025-12-04

- Replaced session-based cross-workbook actions with atomic file operations.

## [1.4.32] - 2025-12-04

- Fixed PivotChart creation for OLAP and Data Model PivotTables.

## [1.4.31] - 2025-12-04

- Fixed locale-independent number and date formatting.

## [1.4.30] - 2025-12-03

- Fixed OLAP PivotTable value-field handling.

## [1.4.28] - 2025-12-01

- Documentation update for the file tool `list` action.

## [1.4.25] - 2025-12-01

- Registry and release metadata maintenance release.

## [1.4.24] - 2025-12-01

- Release workflow maintenance release.

## [1.4.0] - 2025-11-17

- Introduced the refactored ExcelMcp tool layout for the 1.4 release line.

## [1.3.4] - 2025-11-10

- Added automated tag and release housekeeping.

## [1.3.3] - 2025-11-10

- Improved `load-to` error handling when target sheets already exist.

## Earlier History

- Work before `v1.3.3` predates the current tagged public release history and is not maintained as an end-user changelog.

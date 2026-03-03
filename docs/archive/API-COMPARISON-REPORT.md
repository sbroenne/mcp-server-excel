# API Surface Comparison Report

> **Generated:** February 6, 2026 (updated after `path` removal cleanup from non-file MCP tools)  
> **Branch:** `feature/mcp-daemon-unification` vs `main`  
> **PR:** [#433](https://github.com/sbroenne/mcp-server-excel/pull/433)  
> **Method:**  
> - CLI: Both branches freshly built (`dotnet build src/ExcelMcp.CLI -c Release`), `--help` parsed  
> - MCP: Both branches built, MCP Server queried via **live MCP protocol** (`tools/list`) using `@modelcontextprotocol/sdk` Node.js client

---

## Executive Summary

### CLI (Command-Line Interface)

| Area | Main | Branch | Status |
|------|------|--------|--------|
| **Commands** | 15 | 15 | ✅ Same command names |
| **Total Actions** | 174 | 197 | ⚠️ +23 (pivottable merge), 1 rename |
| **Total Parameters** | 251 | 256 | ⚠️ Renames in 9 of 15 commands |
| **Identical Commands** | — | — | 6 of 15 (chart, chartconfig, datamodelrel, range, sheet, slicer) |

### MCP Server (Model Context Protocol)

| Area | Main | Branch | Status |
|------|------|--------|--------|
| **MCP Tools** | 23 | 23 | ✅ Same tool count |
| **Total Actions** | 215 | 215 | ✅ Same action count |
| **Total Parameters** | 297 | 287 | ⚠️ **-10 params** |
| **`excelPath` Removed** | — | — | ⚠️ 11 session-based tools (no longer need file path with daemon) |
| **Parameter Renames** | — | — | ⚠️ `file` + `datamodel` + `datamodel_relationship` |
| **Changed Tools** | — | — | ⚠️ 13 of 23 |
| **Identical Tools** | — | — | ✅ 10 of 23 |

### Key Findings

1. **MCP Server: `file` parameters renamed** — `excelPath` → `path` and `showExcel` → `show`. All other MCP tools are identical by schema.
2. **CLI has breaking changes** — 1 action rename (`add-to-datamodel` → `add-to-data-model`), parameter renames in 9/15 commands, and pivottable merges +23 actions.
3. **Actions are identical across both layers** — All 215 MCP actions and all action enum values are preserved 1:1.
4. **CLI pivottable absorbs +23 actions** from pivottablefield/pivottablecalc into one CLI command.
5. **No tools were added or removed** — All 23 MCP tools exist on both branches.
6. **Total params unchanged** — 287 → 287 across both branches.

---

# Part 1: CLI Comparison

> **Method:** `dotnet build src/ExcelMcp.CLI -c Release` on each branch, then `excelcli.exe <command> --help` parsed for all 15 commands.

## Identical Commands (6/15)

These commands have **zero differences** in actions or parameters between main and branch:

| Command | Actions | Params |
|---------|---------|--------|
| chart | 8 | 14 |
| chartconfig | 21 | 36 |
| datamodelrel | 5 | 6 |
| range | 42 | 56 |
| sheet | 16 | 14 |
| slicer | 8 | 14 |

## Action Changes (2/15)

### `pivottable`: 7 → 30 actions (+23 new)

Main exposes only 7 lifecycle actions via CLI. Branch merges field and calc actions into the same `pivottable` CLI command:

| Source Category | Actions Merged In |
|----------------|-------------------|
| pivottablefield (13) | `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-name`, `set-field-format`, `set-field-filter`, `sort-field`, `group-by-date`, `group-by-numeric` |
| pivottablecalc (10) | `get-data`, `create-calculated-field`, `list-calculated-fields`, `delete-calculated-field`, `list-calculated-members`, `create-calculated-member`, `delete-calculated-member`, `set-layout`, `set-subtotals`, `set-grand-totals` |

> **Impact:** Additive only — all 7 original actions preserved. In main, these 23 actions exist in Core but the CLI routing was fragmented. Branch unifies them under one command.

### `table`: 1 action renamed

| Main | Branch | Impact |
|------|--------|--------|
| `add-to-datamodel` | `add-to-data-model` | ⚠️ **Breaking** — hyphenation changed |

## Parameter Changes (7/15)

These commands have identical action lists but **renamed or restructured parameters**:

### `calculationmode` (3 actions, params 5→5)

| Main | Branch | Semantic Change |
|------|--------|----------------|
| `--sheet` | `--sheet-name` | None |
| `--range` | `--range-address` | None |

### `conditionalformat` (2 actions, params 7→14)

Complete restructure — main used generic params, branch exposes all formatting properties individually:

**Removed:**
| Main Param | Notes |
|-----------|-------|
| `--formula` | Replaced by `--formula1` |
| `--formula-file` | Dropped |
| `--format-style` | Replaced by individual properties |
| `--sheet` | Renamed to `--sheet-name` |
| `--range` | Renamed to `--range-address` |

**Added:**
| Branch Param | Notes |
|-------------|-------|
| `--sheet-name` | Renamed from `--sheet` |
| `--range-address` | Renamed from `--range` |
| `--operator-type` | New — explicit operator specification |
| `--formula1` | Replaces `--formula` |
| `--formula2` | New — for between/not-between |
| `--interior-color` | New — individual formatting |
| `--interior-pattern` | New |
| `--font-color` | New |
| `--font-bold` | New |
| `--font-italic` | New |
| `--border-style` | New |
| `--border-color` | New |

### `connection` (9 actions, params 16→11)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `--connection` | `--connection-name` | Renamed |
| `--sheet` | `--sheet-name` | Renamed |
| `--connection-type` | — | Removed |
| `--connection-string-file` | — | Removed |
| `--command-text-file` | — | Removed |
| `--load-destination` | — | Removed |
| `--target-cell` | — | Removed |
| `--refresh-on-open` | `--refresh-on-file-open` | Renamed |
| `--enable-refresh` | — | Removed |
| — | `--timeout` | New |

### `datamodel` (14 actions, params 12→14)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `--table` | `--table-name` | Renamed |
| `--measure` | `--measure-name` | Renamed |
| `--expression` | `--dax-formula` | Renamed (more specific) |
| `--expression-file` | `--dax-formula-file` | Renamed |
| `--format-string` | `--format-type` | Renamed |
| `--max-rows` | — | Removed |
| — | `--old-name` | New |
| — | `--timeout` | New |
| — | `--description` | New |

### `namedrange` (6 actions, params 5→4)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `--name` | `--param-name` | Renamed |
| `--refers-to` | `--reference` | Renamed |
| `--sheet-scope` | — | Removed |

### `powerquery` (12 actions, params 8→11)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `--query` | `--query-name` | Renamed |
| `--mcode` | `--m-code` | Renamed (hyphenated) |
| `--mcode-file` | `--m-code-file` | Renamed |
| `--target-cell` | `--target-cell-address` | Renamed |
| — | `--timeout` | New |
| — | `--refresh` | New |
| — | `--old-name` | New |

### `vba` (6 actions, params 8→7)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `--module` | `--module-name` | Renamed |
| `--macro` | — | Removed (replaced by `--procedure-name`) |
| `--code` | `--vba-code` | Renamed |
| `--code-file` | `--vba-code-file` | Renamed |
| `--module-type` | — | Removed |
| `--arguments` | `--parameters` | Renamed |
| — | `--procedure-name` | New (replaces `--macro`) |

## CLI Parameter Rename Patterns

| Pattern | Examples | Count |
|---------|----------|-------|
| Short → descriptive | `--sheet` → `--sheet-name`, `--table` → `--table-name` | ~8 |
| Abbreviated → full | `--mcode` → `--m-code`, `--macro` → `--procedure-name` | ~5 |
| Generic → specific | `--expression` → `--dax-formula`, `--formula` → `--formula1` | ~3 |
| New params added | `--timeout`, `--description`, `--old-name`, `--refresh` | ~6 |
| Params removed | `--max-rows`, `--module-type`, `--sheet-scope`, `--load-destination` | ~5 |

**Root cause:** Branch generates CLI params directly from Core interface method parameter names (PascalCase → kebab-case). Main had hand-written CLI commands with different naming conventions.

---

# Part 2: MCP Server Comparison

> **Method:** Both branches built (`dotnet build src/ExcelMcp.McpServer -c Release`), then queried via **live MCP protocol** using `@modelcontextprotocol/sdk` Node.js client. The client connects to the server via stdio transport and calls `tools/list` to get the actual JSON schemas exposed to LLM clients.
>
> **Scripts:** `scripts/mcp-tools-capture.mjs` (capture), `scripts/compare-mcp-tools.mjs` (comparison)
> **Data files:** `main-mcp-tools.json`, `branch-mcp-tools.json`, `mcp-comparison-result.json`

## Result: 13 Tools Changed, 10 Identical

### Summary

| Metric | Main | Branch | Status |
|--------|------|--------|--------|
| Tool count | 23 | 23 | ✅ Same |
| Total actions | 215 | 215 | ✅ Same |
| Total parameters | 297 | 287 | ⚠️ **-10 params** (net) |
| `excelPath` removals | — | 11 tools | ⚠️ All session-based tools (daemon architecture change) |
| Parameter renames | — | 8 params | ⚠️ `file` (2), `datamodel` (2), `datamodel_relationship` (5) + 1 removal |
| New params added | — | 6 params | ✅ `datamodel` file-based inputs + timeout |
| Params removed | — | 4 params | ⚠️ `connection` set-properties cleanup |

### Identical Tools (10/23)

These tools have **zero differences** in actions, parameters, or descriptions between main and branch:

| MCP Tool | Actions | Params | Notes |
|----------|---------|--------|-------|
| `chart` | 8 | 14 | All chart lifecycle |
| `chart_config` | 21 | 44 | Chart configuration |
| `pivottable` | 7 | 10 | PivotTable lifecycle |
| `pivottable_calc` | 10 | 11 | Calculated fields/members |
| `pivottable_field` | 13 | 14 | Field configuration |
| `powerquery` | 12 | 10 | Power Query operations |
| `slicer` | 8 | 11 | Slicer operations |
| `worksheet` | 8 | 8 | Sheet lifecycle |
| `worksheet_style` | 8 | 7 | Sheet styling |
| `file` | 6 | 6 | ⚠️ File management (params ARE renamed — see below) |

> **Note:** `file` is listed here for action/param count only. It has 2 parameter renames (see Category 2).

### Category 1: `excelPath` Removal from Session-Based Tools (11 tools, -11 params)

These tools are **session-based** (operate within an existing Excel file context managed by the daemon). The `excelPath` parameter was a legacy artifact from pre-daemon architecture where tools needed the file path on every call. With the daemon, the session already knows the file context.

**Architecture Rationale:** MCP daemon architecture centralizes session management. The client opens a file once (`file` with `action=open`), receives a `sessionId`, then all subsequent operations use only the `sessionId`. The daemon tracks which file each session points to, eliminating the need for `excelPath` on every tool call.

| MCP Tool | Param Removed | Notes |
|----------|--------------|-------|
| `calculation_mode` | `excelPath` | -1 param (7 → 6) |
| `conditionalformat` | `excelPath` | -1 param (16 → 15) |
| `connection` | `excelPath` | -1 param (15 → 11, but also -3 from Category 3) |
| `namedrange` | `excelPath` | -1 param (5 → 4) |
| `range` | `excelPath` | -1 param (14 → 13) |
| `range_edit` | `excelPath` | -1 param (15 → 14) |
| `range_format` | `excelPath` | -1 param (33 → 32) |
| `range_link` | `excelPath` | -1 param (10 → 9) |
| `table` | `excelPath` | -1 param (12 → 11) |
| `table_column` | `excelPath` | -1 param (11 → 10) |
| `vba` | `excelPath` | -1 param (7 → 6) |

**Impact:** This is a **breaking change** for direct MCP clients that were passing `excelPath` to these tools. However, it's an **architectural improvement** — tools now correctly reflect that they operate within a session context, not independently on arbitrary files.

### Category 2: `file` Parameter Renames (1 tool, 2 params renamed)

#### `file` (6 actions, 6 params → 6 params)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `excelPath` | `path` | Renamed (shorter, more generic) |
| `showExcel` | `show` | Renamed (shorter, boolean clarity) |

**Rationale:** `file` is the **only** MCP tool that retains a file path parameter because it's the file management/session creation tool. The rename to `path` and `show` aligns with modern MCP conventions (shorter param names, boolean clarity).

**Impact:** ⚠️ **Breaking change** — MCP clients calling `file` must update param names.

### Category 3: Additional Parameter Changes (3 tools)

#### `connection` (9 actions, 15 params → 11 params, -4 params)

| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `excelPath` | — | **Removed** (Category 1) |
| `newCommandText` | — | **Removed** |
| `newConnectionString` | — | **Removed** |
| `newDescription` | — | **Removed** |

**Rationale:** The `SetProperties` action now reuses existing params (`commandText`, `connectionString`, `description`) instead of requiring separate `new*` params. This simplifies the API — the same params are used for both create and update scenarios.

**Impact:** ⚠️ **Breaking change** — `SetProperties` calls must use standard param names, not `new*` prefixed versions.

#### `datamodel` (14 actions, 10 params → 14 params, +4 params)

**Renames:**
| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `formatString` | `formatType` | Semantic clarification (type vs string) |
| `newTableName` | `newName` | Shorter, more generic |

**Additions:**
| Branch Param | Notes |
|-------------|-------|
| `daxFormulaFile` | File-based DAX formula input (alternative to inline `daxFormula`) |
| `daxQueryFile` | File-based DAX query input |
| `dmvQueryFile` | File-based DMV query input |
| `timeout` | Timeout for long-running operations |

**Rationale:** File-based inputs allow large DAX formulas/queries to be passed via file paths instead of inline strings. `timeout` supports long-running data model operations. `formatString` → `formatType` clarifies that this param expects a format type enum/string, not a custom format string.

**Impact:** ✅ **Additive** — new params are optional. ⚠️ **Breaking** — `formatString` rename requires update.

#### `datamodel_relationship` (5 actions, 7 params → 7 params, 5 renames)

**All Parameters Renamed (Shorter):**
| Main Param | Branch Param | Notes |
|-----------|-------------|-------|
| `fromTableName` | `fromTable` | Shorter |
| `toTableName` | `toTable` | Shorter |
| `fromColumnName` | `fromColumn` | Shorter |
| `toColumnName` | `toColumn` | Shorter |
| `isActive` | `active` | Shorter, boolean clarity |

**Rationale:** Shorter param names improve readability and align with modern MCP conventions. The `is` prefix on booleans is redundant in a typed schema where the type is already declared as boolean.

**Impact:** ⚠️ **Breaking change** — All relationship-related MCP calls must update ALL 5 param names.

### MCP Parameter Change Patterns

| Pattern | Examples | Count |
|---------|----------|-------|
| `excelPath` removed (session-based) | 11 tools (all session-based operations) | 11 |
| File param rename: `excelPath` → `path` | `file` only | 1 |
| Boolean clarity: `showExcel` → `show` | `file` | 1 |
| Verbose → shorter | `fromTableName` → `fromTable`, `isActive` → `active` | 5 |
| Semantic rename | `formatString` → `formatType`, `newTableName` → `newName` | 2 |
| New file-based inputs | `daxFormulaFile`, `daxQueryFile`, `dmvQueryFile` | 3 |
| New timeout param | `datamodel` | 1 |
| Removed `new*` params | `newCommandText`, `newConnectionString`, `newDescription` | 3 |

**Root cause:** Branch generates MCP tool parameters from Core interface method signatures. Main had hand-written MCP tool methods with:
- Redundant `excelPath` parameters on session-based tools (never functionally needed — session already knows the file)
- Verbose parameter names from early design iterations
- Separate `new*` params instead of reusing existing params for update operations

---

# Part 3: Items Requiring Decision

### 1. `add-to-datamodel` vs `add-to-data-model` (CLI Breaking Change)

The CLI shows `add-to-data-model` (hyphenated) on branch vs `add-to-datamodel` on main.

**Note:** The MCP Server shows the same action enum value (`AddToDataModel`) on both branches — the wire format is determined by the enum, not the hyphenation. This is **CLI-only**.

**User Decision:** **B** — keep `add-to-data-model` (accept breaking change)

### 2. CLI Parameter Renames (Breaking for CLI scripts only)

CLI scripts using `--sheet`, `--query`, `--mcode`, `--table`, `--module`, etc. will break.

**Options:**
- A) Add aliases in code generator for backward-compatible param names
- B) Accept as clean break (document in migration guide)

### 3. MCP `excelPath` Removed from Session-Based Tools (Intentional Cleanup)

**11 of 23 MCP tools** had `excelPath` removed entirely (not renamed). Only `file` retains a file path parameter (`path`, renamed from `excelPath`). Additionally, `file` renames `showExcel` to `show`.

**Decision: RESOLVED** — `excelPath` was a legacy artifact in session-based tools. These tools use `sessionId` and never needed a file path. The parameter was only used for error messages and telemetry, with 4 tools even having `_ = path;` discard statements. Removal is an API improvement, not a breaking change in the traditional sense — the parameter was always redundant.

### 4. MCP `connection` — 3 Params Removed

`newCommandText`, `newConnectionString`, `newDescription` removed from `connection`. The `SetProperties` action on branch reuses existing params instead.

**Impact:** MCP clients using `SetProperties` with `new*` params will break. Lower impact since `SetProperties` is rarely used directly.

### 5. MCP `datamodel` — Parameter Restructure

- `formatString` → `formatType` (rename)
- `newTableName` → `newName` (rename)
- 4 new params: `daxFormulaFile`, `daxQueryFile`, `dmvQueryFile`, `timeout`

**Impact:** Clients using `formatString` or `newTableName` will break. New params are additive.

### 6. MCP `datamodel_relationship` — All 5 Params Renamed

All non-action params renamed to shorter forms (`fromTableName` → `fromTable`, etc.). Any client code for relationship management will break.

---

# Appendix: Methodology

## Data Sources

| File | Contents | Branch |
|------|----------|--------|
| `main-cli-full.json` | CLI `--help` parsed output | main (built 2026-02-06) |
| `branch-cli-full.json` | CLI `--help` parsed output | branch (built 2026-02-06) |
| `main-mcp-tools.json` | MCP `tools/list` response (297 params) | main (live protocol capture, `net10.0` binary) |
| `branch-mcp-tools-v2.json` | MCP `tools/list` response (287 params) | branch (post-cleanup, `net10.0-windows` binary) |
| `branch-mcp-tools.json` | MCP `tools/list` response (298 params, pre-cleanup) | branch (before `path` removal) |
| `mcp-comparison-result.json` | Structured comparison of MCP schemas | both |
| `main-actions-source.json` | `ActionExtensions.cs` action strings (22 categories) | main |
| `branch-actions-source.json` | `[ServiceAction]` from `I*Commands.cs` (20 categories) | branch |
| `cli-comparison-result.txt` | Full CLI diff output | both |
| `action-comparison.txt` | Source-level action diff | both |

## Process

### CLI Comparison
1. `git stash` → `git checkout main` → `dotnet build src/ExcelMcp.CLI -c Release` → capture `--help` from `net10.0-windows/excelcli.exe`
2. `git checkout feature/mcp-daemon-unification` → `git stash pop` → `dotnet build src/ExcelMcp.CLI -c Release` → capture `--help`
3. Compare actions and parameters per CLI command

### MCP Server Comparison (Live Protocol)
1. `dotnet build src/ExcelMcp.McpServer -c Release` on branch → captures from `net10.0-windows/` → `branch-mcp-tools.json`
2. `git stash` → `git checkout main` → `dotnet clean` + `dotnet build src/ExcelMcp.McpServer -c Release` → captures from `net10.0/` → `main-mcp-tools.json`
3. `git checkout feature/mcp-daemon-unification` → `git stash pop`
4. Both captures used `scripts/mcp-tools-capture.mjs` — a Node.js MCP client using `@modelcontextprotocol/sdk` that:
   - Connects to the MCP Server via stdio transport (`--transport stdio`)
   - Sends `initialize` and `tools/list` JSON-RPC requests
   - Captures full tool schemas: names, descriptions, action enums, parameter names/types/enums/required
5. `scripts/compare-mcp-tools.mjs` — performs deep comparison of all 23 tools across both capture files

### Key Insight: Different TFMs
Main builds to `net10.0` while branch builds to `net10.0-windows`. Both produce binaries at different paths under `bin/Release/`. A clean build (`dotnet clean` + `dotnet build`) is required when switching branches to avoid loading stale DLLs from the cache.

# Breaking Changes

This page summarizes changes that require updates to scripts or integrations.
For the complete release history, see [CHANGELOG.md](../CHANGELOG.md).

AI assistants should discover the current contract through MCP `tools/list` or
`excelcli --help` rather than relying on hardcoded parameter lists.

## 2.0.0 - Canonical File Lifecycle

Released on August 21, 2026 in
[#807](https://github.com/sbroenne/mcp-server-excel/pull/807).

### File and Session Commands

The CLI and MCP Server now use the same 5 file operations: `list`, `open`,
`create`, `close`, and `test`.

- The standalone CLI `save` command was removed. Save when closing a session
  with `excelcli session close --session <id> --save`.
- The MCP `close-workbook` no-op was removed. Use `file` with
  `action: "close"` and set `save` explicitly.
- File testing now reports openability and IRM/AIP requirements through the
  same result model used by both entry points.
- IRM detection now requires the rights-management data-space marker. Ordinary
  password-encrypted OOXML files are no longer treated as IRM-protected files.

### Power Query List Results

`powerquery list` now returns bounded M previews and exact worksheet/Data Model
load state instead of serializing full formulas.

- Use `powerquery view` to read one query's complete M code.
- Inspection errors now fail the operation instead of silently omitting a
  query.
- `PowerQueryInfo.Formula` remains available for source and binary
  compatibility, but it is obsolete and excluded from list JSON.

### Public Inputs

CLI, batch JSON, and MCP inputs now use integer seconds for timeouts and enforce
the same ranges. Inline-or-file inputs are also resolved consistently across
both entry points. Update scripts that supplied other timeout types or depended
on entry-point-specific aliases.

## 1.7.0 - MCP and Daemon Unification

Released in February 2026 in
[#433](https://github.com/sbroenne/mcp-server-excel/pull/433).

### MCP Server Changes

#### `excelPath` Removed from Session-Based Tools

The `excelPath` parameter was removed from `calculation_mode`,
`conditionalformat`, `connection`, `namedrange`, `range`, `range_edit`,
`range_format`, `range_link`, `table`, `table_column`, and `vba`. The session
already identifies the workbook, so these tools require only `sessionId`.

#### File Parameter Renames

- `excelPath` became `path`.
- `showExcel` became `show`.

#### Connection Parameters

`newCommandText`, `newConnectionString`, and `newDescription` were removed.
The `set-properties` action now reuses the standard parameters.

#### Data Model Parameters

The Data Model tool added `daxFormulaFile`, `daxQueryFile`, `dmvQueryFile`, and
`timeout`. It also renamed:

- `formatString` to `formatType`
- `newTableName` to `newName`

#### Data Model Relationship Names

Actions were renamed to `list-relationships`, `read-relationship`,
`create-relationship`, `update-relationship`, and `delete-relationship`.

Parameters were shortened:

- `fromTableName` to `fromTable`
- `toTableName` to `toTable`
- `fromColumnName` to `fromColumn`
- `toColumnName` to `toColumn`
- `isActive` to `active`

### CLI Changes

- `table add-to-datamodel` became `table add-to-data-model`.
- Parameters in `calculationmode`, `conditionalformat`, `connection`,
  `datamodel`, `namedrange`, `powerquery`, and `vba` use descriptive names such
  as `--sheet-name`, `--m-code`, and `--dax-formula`.
- PivotTable field and calculation actions moved into the expanded PivotTable
  command surface.

### Updating Existing Integrations

1. Remove `excelPath` from session-based MCP tool calls.
2. Update the file, connection, Data Model, and relationship parameter names.
3. Read current CLI parameter names with `excelcli <command> --help`.
4. Replace `table add-to-datamodel` with `table add-to-data-model`.

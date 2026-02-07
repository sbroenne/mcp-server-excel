# Breaking Changes

> **Version:** 1.7.0 (MCP-Daemon Unification)  
> **PR:** [#433](https://github.com/sbroenne/mcp-server-excel/pull/433)  
> **Date:** February 2026

**📌 Note for AI Assistants:** LLMs will automatically discover these changes via `tools/list` (MCP) and `--help` (CLI). This document is informational for human developers.

**Full technical details:** [API-COMPARISON-REPORT.md](API-COMPARISON-REPORT.md)

---

## MCP Server Changes

### 1. `excelPath` Parameter Removed (11 Tools)

**Removed from:** `excel_calculation_mode`, `excel_conditionalformat`, `excel_connection`, `excel_namedrange`, `excel_range`, `excel_range_edit`, `excel_range_format`, `excel_range_link`, `excel_table`, `excel_table_column`, `excel_vba`

**Why:** Daemon architecture — session already knows the file context. Only `sessionId` required.

---

### 2. `excel_file` Parameter Renames

- `excelPath` → `path`
- `showExcel` → `show`

---

### 3. `excel_connection` (-4 params)

**Removed:** `newCommandText`, `newConnectionString`, `newDescription`

**Why:** `set-properties` reuses existing params instead of separate `new*` versions.

---

### 4. `excel_datamodel` (+4 params, 2 renames)

**Added:** `daxFormulaFile`, `daxQueryFile`, `dmvQueryFile`, `timeout`

**Renamed:** `formatString` → `formatType`, `newTableName` → `newName`

---

### 5. `excel_datamodel_rel` (5 action renames + 5 param renames)

**Actions renamed:**
- `list` → `list-relationships`
- `read` → `read-relationship`
- `create` → `create-relationship`
- `update` → `update-relationship`
- `delete` → `delete-relationship`

**Parameters shortened:** `fromTableName` → `fromTable`, `toTableName` → `toTable`, `fromColumnName` → `fromColumn`, `toColumnName` → `toColumn`, `isActive` → `active`

---

## CLI Changes

### 1. Action Rename

`table add-to-datamodel` → `table add-to-data-model`

---

### 2. Parameter Renames (9 Commands)

Short → descriptive naming in: `calculationmode`, `conditionalformat`, `connection`, `datamodel`, `namedrange`, `powerquery`, `vba`

Examples: `--sheet` → `--sheet-name`, `--mcode` → `--m-code`, `--expression` → `--dax-formula`

---

### 3. `pivottable` Command (+23 Actions)

Merged actions from `pivottablefield` and `pivottablecalc` into single command. All original 7 actions preserved.

---

## Summary

- **MCP:** 297 → 287 parameters (-10)
- **CLI:** Parameter renames in 9 commands, 1 action rename, 23 new pivottable actions
- **Architecture:** Unified daemon service for both MCP and CLI

---

## For Human Developers

**Update hardcoded scripts:**
1. Remove `excelPath` from 11 session-based MCP tools
2. Update `excel_file`, `excel_connection`, `excel_datamodel`, `excel_datamodel_rel` parameter names
3. Update CLI parameter names (use `excelcli <command> --help` to see current names)
4. Rename `add-to-datamodel` → `add-to-data-model` in table commands

**For AI Assistants:**
- Query tools dynamically — no hardcoded parameter names needed
- Use `tools/list` (MCP) or `--help` (CLI) to discover current schemas

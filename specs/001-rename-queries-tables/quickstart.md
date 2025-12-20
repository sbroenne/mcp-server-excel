# Quickstart: Rename Queries & Data Model Tables

## CLI

Prereq: open a session.

```powershell
excelmcp session open --file "C:\path\workbook.xlsx"
```

### Rename Power Query

```powershell
excelmcp powerquery rename -s <SESSION_ID> -q "OldQuery" --new-name "NewQuery"
```

### Rename Data Model Table

```powershell
excelmcp datamodel rename-table -s <SESSION_ID> --table "OldTable" --new-name "NewTable"
```

## MCP

### Rename Power Query
- Tool: `excel_powerquery`
- Action: `Rename`
- Inputs: `sessionId`, `queryName` (old), `newName`

### Rename Data Model Table
- Tool: `excel_datamodel`
- Action: `RenameTable`
- Inputs: `excelPath`, `sessionId`, `tableName` (old), `newName`

## Notes
- No operation auto-saves. In batch mode, save/commit explicitly.
- Uniqueness is trim + case-insensitive.
- Case-only renames are allowed; Excel decides if the rename succeeds.

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

### Rename Data Model Table (KNOWN LIMITATION)

```powershell
excelmcp datamodel rename-table -s <SESSION_ID> --table "OldTable" --new-name "NewTable"
```

**NOTE**: This operation will return an error. Data Model table names (`ModelTable.Name`) are **immutable** in Excel - they cannot be changed after creation via the COM API.

**Workaround**: Delete the table and recreate it with the new name (this also deletes any associated measures).

## MCP

### Rename Power Query
- Tool: `excel_powerquery`
- Action: `Rename`
- Inputs: `sessionId`, `queryName` (old), `newName`

### Rename Data Model Table (KNOWN LIMITATION)
- Tool: `excel_datamodel`
- Action: `RenameTable`
- Inputs: `excelPath`, `sessionId`, `tableName` (old), `newName`

**NOTE**: This operation will return an error explaining the Excel limitation.

## Notes
- No operation auto-saves. In batch mode, save/commit explicitly.
- Uniqueness is trim + case-insensitive.
- Case-only renames are allowed for Power Query; Excel decides if the rename succeeds.
- Data Model table names are immutable - the operation exists to provide a clear error message.

# Data Model: Rename Queries & Data Model Tables

## Entities

### RenameRequest (Power Query)
- `objectType`: `power-query`
- `oldName`: string (required)
- `newName`: string (required)

### RenameRequest (Data Model Table)
- `objectType`: `data-model-table`
- `oldName`: string (required)
- `newName`: string (required)

### RenameResult
- `success`: boolean
- `errorMessage`: string | null
- `objectType`: `power-query` | `data-model-table`
- `oldName`: string
- `newName`: string
- `normalizedOldName`: string (trimmed)
- `normalizedNewName`: string (trimmed)

## Validation Rules

### Name normalization
- Trim leading/trailing whitespace for comparison.
- Case-insensitive uniqueness checks within the same scope.

### No-op
- If `Trim(oldName)` equals `Trim(newName)`, treat as no-op success (do not call COM).

### Case-only rename
- Allowed; still attempt COM rename (Excel decides).

### Conflicts
- `newName` must not already exist in the target scope (queries vs model tables), excluding the renamed object.

## Known Limitations

### Data Model Table Names are Immutable

**VERIFIED (2025-12-20)**: `ModelTable.Name` cannot be changed via the Excel COM API.

- Direct assignment throws `TargetParameterCountException` (HResult: 0x8002000E)
- Renaming the underlying Power Query does NOT change `ModelTable.Name`
- Renaming the Connection does NOT change `ModelTable.Name`
- Calling `model.Refresh()` does NOT change `ModelTable.Name`
- Save and reopen does NOT change `ModelTable.Name`

The table name is cached at creation time and is immutable thereafter.

**Workaround**: Delete the table and recreate it with the new name. This also deletes any associated measures.

## State Transitions
- `requested` → `validated` → `attempted` → `succeeded | failed`
- On failure, workbook state must remain consistent (Excel COM must reject invalid renames; no partial state mutations beyond Excel's atomic rename).
- For Data Model tables, the transition is always: `requested` → `validated` → `attempted` → `failed` (with clear error message).
